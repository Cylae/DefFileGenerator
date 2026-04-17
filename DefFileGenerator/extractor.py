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
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    from defusedxml import ElementTree as ET
    HAS_DEFUSEDXML = True
except ImportError:
    HAS_DEFUSEDXML = False

from DefFileGenerator.def_gen import Generator

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
        'Length': ['length', 'len', 'size', 'count', 'quantity'],
        'StartBit': ['startbit', 'bit offset', 'bit', 'start']
    }

    TYPE_PATTERN = re.compile(r'^(u|i|uint|int)(\d+)$', re.IGNORECASE)

    def __init__(self, mapping=None):
        self.mapping = mapping or {}
        self.type_mapping = {
            'uint16': 'U16', 'int16': 'I16', 'uint32': 'U32', 'int32': 'I32',
            'float32': 'F32', 'float': 'F32', 'u16': 'U16', 'i16': 'I16',
            'u32': 'U32', 'i32': 'I32', 'f32': 'F32', 'string': 'STRING', 'bits': 'BITS'
        }


    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb[sheet_name] if sheet_name else wb.active
        data = []
        rows = list(ws.rows)
        if not rows: return []
        headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]
        for row in rows[1:]:
            data.append({headers[i]: cell.value for i, cell in enumerate(row) if i < len(headers)})
        return data

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        data = []
        try:
            with pdfplumber.open(filepath) as pdf:
                target_pages = pdf.pages if pages is None else [pdf.pages[i-1] for i in (pages if isinstance(pages, list) else [pages])]
                for page in target_pages:
                    tables = page.extract_tables()
                    logging.debug(f"Found {len(tables)} tables on page {page.page_number}")
                    for table in tables:
                        if not table or len(table) < 2: continue
                        headers = [str(c).replace('\n', ' ').strip() if c else "" for c in table[0]]
                        for row in table[1:]:
                            row_dict = {}
                            for i, cell in enumerate(row):
                                if i < len(headers):
                                    row_dict[headers[i]] = str(cell).replace('\n', ' ').strip() if cell else ""
                            data.append(row_dict)
        except Exception as e:
            logging.error(f"Error extracting from PDF {filepath}: {e}")
        return data

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
                return list(reader)
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

            # Simple conversion of XML nodes to list of dicts
            data = []
            for child in root:
                row = {}
                for subchild in child:
                    row[subchild.tag] = subchild.text
                if row:
                    data.append(row)
            return data
        except Exception as e:
            logging.error(f"Error extracting from XML: {e}")
            return []

    def map_and_clean(self, tables):
        if not tables: return []
        # Support single table (list of dicts) or list of tables
        if isinstance(tables, list) and tables and isinstance(tables[0], dict):
            tables = [tables]

        final_data = []
        generator = Generator()

        for table in tables:
            if not table: continue

            # Robust column detection: scan first 5 rows
            col_map = {}
            used_src_cols = set()
            all_src_cols = set()
            for row in table[:5]:
                all_src_cols.update(row.keys())

            # 1. Explicit mapping
            for target, source in self.mapping.items():
                if any(source in row for row in table[:5]):
                    col_map[target] = source
                    used_src_cols.add(source)

            # 2. Priority fuzzy matching
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag', 'Factor', 'ScaleFactor', 'Length', 'StartBit']
            for target in detection_order:
                if target in col_map: continue
                patterns = self.COLUMN_MAPPING.get(target, [target.lower()])

                found = False
                for src_col in all_src_cols:
                    if src_col in used_src_cols: continue
                    if any(p in str(src_col).lower() for p in patterns):
                        col_map[target] = src_col
                        used_src_cols.add(src_col)
                        found = True
                        break
                if found: continue

            for row in table:
                new_row = {target: row.get(src_col) for target, src_col in col_map.items()}
                if not new_row.get('Name') and not new_row.get('Address'): continue

                # Normalize Type
                new_row['Type'] = Generator.normalize_type(new_row.get('Type', 'U16'))

                # Normalize Address
                addr = str(new_row.get('Address', '')).strip()
                start_bit = str(new_row.get('StartBit', '')).strip()
                length = str(new_row.get('Length', '')).strip()

                if '_' in addr:
                    parts = addr.split('_')
                    new_row['Address'] = '_'.join(Generator.normalize_address_val(p) for p in parts)
                else:
                    base_addr = Generator.normalize_address_val(addr)
                    if new_row['Type'] == 'BITS' and start_bit:
                        if not length: length = '1'
                        new_row['Address'] = f"{base_addr}_{Generator.normalize_address_val(start_bit)}_{Generator.normalize_address_val(length)}"
                    elif length and new_row['Type'] == 'STRING':
                        new_row['Address'] = f"{base_addr}_{Generator.normalize_address_val(length)}"
                    else:
                        new_row['Address'] = base_addr

                # Normalize Factor (fractions like 1/10)
                factor = str(new_row.get('Factor', '1'))
                if '/' in factor:
                    try:
                        p = factor.split('/')
                        new_row['Factor'] = str(float(p[0]) / float(p[1]))
                    except (ValueError, ZeroDivisionError):
                        new_row['Factor'] = '1'

                if 'RegisterType' not in new_row:
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

    if ext in ['.xlsx', '.xlsm']: raw = extractor.extract_from_excel(args.input_file, args.sheet)
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
