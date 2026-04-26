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
    import defusedxml
    from defusedxml import ElementTree as ET
    HAS_DEFUSEDXML = True
except ImportError:
    defusedxml = None
    HAS_DEFUSEDXML = False

try:
    from DefFileGenerator.def_gen import Generator
except ImportError:
    # Support local import if running from within the directory
    try:
        from def_gen import Generator
    except ImportError:
        Generator = None

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
        'Offset': ['offset', 'bias', 'coefficient b'],
        'ScaleFactor': ['scalefactor', 'scale factor'],
        'Length': ['length', 'len', 'size', 'count', 'quantity'],
        'StartBit': ['startbit', 'bit offset', 'bit', 'start']
    }

    def __init__(self, mapping=None):
        self.mapping = mapping or {}

    @staticmethod
    def normalize_type(t):
        if Generator:
            return Generator.normalize_type(t)
        return str(t).upper() if t else 'U16'

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            all_tables = []
            sheets = [wb[sheet_name]] if sheet_name else wb.worksheets
            for ws in sheets:
                data = []
                rows = list(ws.rows)
                if not rows: continue
                headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]
                for row in rows[1:]:
                    data.append({headers[i]: cell.value for i, cell in enumerate(row) if i < len(headers)})
                if data:
                    all_tables.append(data)
            return all_tables
        except Exception as e:
            logging.error(f"Error extracting from Excel {filepath}: {e}")
            return []

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
                    logging.debug(f"Found {len(tables)} tables on page {page.page_number}")
                    for table in tables:
                        if not table or len(table) < 2: continue
                        table_data = []
                        headers = [str(c).replace('\n', ' ').strip() if c else "" for c in table[0]]
                        for row in table[1:]:
                            row_dict = {}
                            for i, cell in enumerate(row):
                                if i < len(headers):
                                    row_dict[headers[i]] = str(cell).replace('\n', ' ').strip() if cell else ""
                            table_data.append(row_dict)
                        if table_data:
                            all_tables.append(table_data)
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
                return [list(reader)]
        except Exception as e:
            logging.error(f"Error extracting from CSV {filepath}: {e}")
            return []

    def extract_from_xml(self, filepath):
        if not HAS_DEFUSEDXML:
            logging.error("defusedxml is required for secure XML parsing.")
            return []
        try:
            with open(filepath, 'rb') as f:
                tree = ET.parse(f)
                root = tree.getroot()

            data = []
            for elem in root.iter():
                row = {}
                for child in elem:
                    if len(child) == 0 and child.text:
                        row[child.tag] = child.text.strip()
                if len(row) >= 2:
                    data.append(row)

            unique_data = []
            seen = set()
            for d in data:
                js = json.dumps(d, sort_keys=True)
                if js not in seen:
                    unique_data.append(d)
                    seen.add(js)

            return [unique_data] if unique_data else []
        except Exception as e:
            # Allow security exceptions to propagate for robust testing
            if HAS_DEFUSEDXML and isinstance(e, (defusedxml.EntitiesForbidden, defusedxml.ExternalReferenceForbidden, defusedxml.DTDForbidden)):
                raise
            logging.error(f"Error extracting from XML {filepath}: {e}")
            return []

    def map_and_clean(self, tables, address_offset=0):
        if not tables: return []
        # Support single table (list of dicts) or list of tables
        if isinstance(tables, list) and tables and isinstance(tables[0], dict):
            tables = [tables]

        final_data = []
        generator = Generator() if Generator else None

        for table in tables:
            if not table: continue

            all_keys = set()
            for row in table[:5]:
                all_keys.update(row.keys())

            col_map = {}
            used_src_cols = set()

            # 1. Explicit mapping
            for target, source in self.mapping.items():
                if source in all_keys:
                    col_map[target] = source
                    used_src_cols.add(source)

            # 2. Priority fuzzy matching
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag', 'Factor', 'Offset', 'ScaleFactor', 'Length', 'StartBit']
            for target in detection_order:
                if target in col_map: continue
                patterns = self.COLUMN_MAPPING.get(target, [target.lower()])
                for src_col in all_keys:
                    if src_col in used_src_cols: continue
                    if any(p in str(src_col).lower() for p in patterns):
                        col_map[target] = src_col
                        used_src_cols.add(src_col)
                        break

            for row in table:
                new_row = {target: row.get(src_col) for target, src_col in col_map.items()}
                if not new_row.get('Name') and not new_row.get('Address'): continue

                # Extract StartBit and Length
                sbit = row.get(col_map.get('StartBit'))
                slen = row.get(col_map.get('Length'))
                sbit = str(sbit).strip() if sbit is not None else ''
                slen = str(slen).strip() if slen is not None else ''

                # Normalize Type
                dtype = self.normalize_type(new_row.get('Type', 'U16'))
                new_row['Type'] = dtype

                # Address normalization/construction
                addr = str(new_row.get('Address', '')).strip()
                if dtype == 'BITS':
                    if sbit != '':
                        if slen == '': slen = '1'
                        base_addr = addr.split('_')[0]
                        addr = f"{base_addr}_{sbit}_{slen}"
                    elif '_' not in addr:
                        # Default to StartBit 0 and Length 1 if not already a compound address
                        addr = f"{addr}_0_1"

                if generator:
                    new_row['Address'] = generator.apply_address_offset(addr, address_offset)
                else:
                    new_row['Address'] = addr

                # Factor
                if new_row.get('Factor') is not None:
                    if Generator:
                        new_row['Factor'] = str(Generator._parse_numeric(new_row['Factor'], 1.0))

                if 'RegisterType' not in new_row or not new_row['RegisterType']:
                    new_row['RegisterType'] = 'Holding Register'

                final_data.append(new_row)
        return final_data

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(description='Extract register information.')
    parser.add_argument('input_file'); parser.add_argument('-o', '--output')
    parser.add_argument('--mapping'); parser.add_argument('--sheet'); parser.add_argument('--pages')
    parser.add_argument('--address-offset', type=int, default=0)
    args = parser.parse_args()

    mapping = {}
    if args.mapping:
        with open(args.mapping, 'r') as f: mapping = json.load(f)

    extractor = Extractor(mapping)
    ext = os.path.splitext(args.input_file)[1].lower()

    pages = None
    if args.pages:
        try:
            pages = [int(p.strip()) for p in args.pages.split(',')]
        except ValueError:
            logging.error("Invalid format for --pages. Expected comma-separated integers.")
            sys.exit(1)

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']: raw = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf': raw = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv': raw = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml': raw = extractor.extract_from_xml(args.input_file)
    else: logging.error(f"Unsupported extension: {ext}"); sys.exit(1)

    mapped = extractor.map_and_clean(raw, args.address_offset)
    out = open(args.output, 'w', newline='', encoding='utf-8') if args.output else sys.stdout
    writer = csv.DictWriter(out, fieldnames=['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor'], extrasaction='ignore')
    writer.writeheader(); writer.writerows(mapped)
    if args.output: out.close()

if __name__ == "__main__":
    main()
