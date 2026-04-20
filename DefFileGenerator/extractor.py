#!/usr/bin/env python3
import argparse
import csv
import json
import logging
import os
import re
import sys
import io

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    from defusedxml import ElementTree as ET
    HAS_DEFUSEDXML = True
except ImportError:
    HAS_DEFUSEDXML = False

try:
    from DefFileGenerator.def_gen import Generator
except ImportError:
    from def_gen import Generator

class Extractor:
    COLUMN_MAPPING = {
        'RegisterType': ['register type', 'reg type', 'modbus type', 'registertype'],
        'Address': ['address', 'addr', 'offset', 'register', 'reg'],
        'Name': ['name', 'description', 'parameter', 'variable', 'signal', 'signal name'],
        'Type': ['data type', 'datatype', 'type', 'format'],
        'Unit': ['unit', 'units'],
        'Tag': ['tag'],
        'Action': ['action', 'access'],
        'Factor': ['scale', 'factor', 'multiplier', 'ratio'],
        'ScaleFactor': ['scalefactor'],
        'Offset': ['offset', 'bias', 'coefficient b'],
        'Length': ['length', 'len', 'size', 'count', 'quantity'],
        'StartBit': ['startbit', 'bit offset', 'bit', 'start']
    }

    def __init__(self, mapping=None):
        self.mapping = mapping or {}

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
        all_tables = []
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            sheets = [sheet_name] if sheet_name else wb.sheetnames
            for s_name in sheets:
                ws = wb[s_name]
                data = []
                rows = list(ws.rows)
                if not rows: continue
                headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]
                for row in rows[1:]:
                    data.append({headers[i]: cell.value for i, cell in enumerate(row) if i < len(headers)})
                if data:
                    all_tables.append(data)
        except Exception as e:
            logging.error(f"Error extracting from Excel {filepath}: {e}")
        return all_tables

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        all_tables = []
        try:
            with pdfplumber.open(filepath) as pdf:
                target_pages = pdf.pages if pages is None else [pdf.pages[i-1] for i in (pages if isinstance(pages, list) else [pages])]
                for page in target_pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if not table or len(table) < 2: continue
                        headers = [str(c).replace('\n', ' ').strip() if c else "" for c in table[0]]
                        data = []
                        for row in table[1:]:
                            row_dict = {}
                            for i, cell in enumerate(row):
                                if i < len(headers):
                                    row_dict[headers[i]] = str(cell).replace('\n', ' ').strip() if cell else ""
                            data.append(row_dict)
                        if data:
                            all_tables.append(data)
        except Exception as e:
            logging.error(f"Error extracting from PDF {filepath}: {e}")
        return all_tables

    def extract_from_csv(self, filepath):
        try:
            with open(filepath, 'rb') as f:
                content = f.read()
                encoding = 'utf-16' if content.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'
            with open(filepath, 'r', encoding=encoding) as f:
                snippet = f.read(1024)
                f.seek(0)
                delimiter = ','
                for d in [',', ';', '\t']:
                    if d in snippet:
                        delimiter = d; break
                reader = csv.DictReader(f, delimiter=delimiter)
                data = list(reader)
                return [data] if data else []
        except Exception as e:
            logging.error(f"Error extracting from CSV: {e}")
            return []

    def extract_from_xml(self, filepath):
        if not HAS_DEFUSEDXML:
            logging.error("defusedxml is required for secure XML parsing.")
            return []
        try:
            with open(filepath, 'rb') as f:
                tree = ET.parse(f)
                root = tree.getroot()
            all_data = []
            for child in root:
                row_data = {}
                for subchild in child:
                    if len(subchild) == 0:
                        row_data[subchild.tag] = subchild.text
                if row_data:
                    all_data.append(row_data)
            return [all_data] if all_data else []
        except Exception as e:
            logging.error(f"Error extracting from XML: {e}")
            return []

    def map_and_clean(self, tables):
        if not tables: return []
        # Support single table (list of dicts) or list of tables
        if isinstance(tables, list) and tables and isinstance(tables[0], dict):
            tables = [tables]

        final_data = []
        targets = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag', 'Factor', 'ScaleFactor', 'Offset', 'Length', 'StartBit']

        for table in tables:
            if not table: continue
            col_map = {}
            used_src_cols = set()

            # Priority 1: Explicit mapping
            first_row = table[0]
            for target, source in self.mapping.items():
                if source in first_row:
                    col_map[target] = source
                    used_src_cols.add(source)

            # Priority 2: Fuzzy matching (scan first 5 rows)
            rows_to_scan = table[:5]
            for target in targets:
                if target in col_map: continue
                patterns = self.COLUMN_MAPPING.get(target, [target.lower()])
                found = False
                for row in rows_to_scan:
                    for src_col in row.keys():
                        if src_col in used_src_cols: continue
                        if any(p in str(src_col).lower() for p in patterns):
                            col_map[target] = src_col
                            used_src_cols.add(src_col)
                            found = True
                            break
                    if found: break

            for row in table:
                new_row = {target: row.get(src_col) for target, src_col in col_map.items()}
                if not new_row.get('Name') and not new_row.get('Address'): continue

                raw_addr = new_row.get('Address')
                addr = str(raw_addr).strip() if raw_addr is not None else ""

                if '_' in addr:
                    addr = '_'.join(Generator.normalize_address_val(p) for p in addr.split('_'))
                else:
                    addr = Generator.normalize_address_val(addr)

                start_bit_val = new_row.get('StartBit')
                start_bit = str(start_bit_val).strip() if start_bit_val is not None else ""

                length_val = new_row.get('Length')
                length = str(length_val).strip() if length_val is not None else ""

                dtype = Generator.normalize_type(new_row.get('Type', 'U16'))
                new_row['Type'] = dtype

                if start_bit:
                    if not length and dtype == 'BITS':
                        length = '1'
                    if addr:
                        addr = f"{addr}_{start_bit}" + (f"_{length}" if length else "")
                elif length and dtype == 'STRING' and '_' not in addr:
                    if addr:
                        addr = f"{addr}_{length}"

                new_row['Address'] = addr

                factor_val = new_row.get('Factor')
                factor = str(factor_val).strip() if factor_val is not None else "1"
                if '/' in factor:
                    try:
                        p = factor.split('/')
                        new_row['Factor'] = str(float(p[0]) / float(p[1]))
                    except (ValueError, ZeroDivisionError):
                        new_row['Factor'] = '1'

                if not new_row.get('RegisterType'):
                    new_row['RegisterType'] = 'Holding Register'

                final_data.append(new_row)
        return final_data

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(description='Extract register information.')
    parser.add_argument('input_file'); parser.add_argument('-o', '--output')
    parser.add_argument('--mapping'); parser.add_argument('--sheet'); parser.add_argument('--pages')
    args = parser.parse_args()

    mapping = {}
    if args.mapping:
        with open(args.mapping, 'r') as f: mapping = json.load(f)

    extractor = Extractor(mapping)
    ext = os.path.splitext(args.input_file)[1].lower()

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']: raw = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf': raw = extractor.extract_from_pdf(args.input_file, [int(p) for p in args.pages.split(',')] if args.pages else None)
    elif ext == '.csv': raw = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml': raw = extractor.extract_from_xml(args.input_file)
    else: logging.error(f"Unsupported extension: {ext}"); sys.exit(1)

    mapped = extractor.map_and_clean(raw)
    out = open(args.output, 'w', newline='', encoding='utf-8') if args.output else sys.stdout
    writer = csv.DictWriter(out, fieldnames=['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor'], extrasaction='ignore')
    writer.writeheader(); writer.writerows(mapped)
    if args.output: out.close()

if __name__ == "__main__":
    main()
