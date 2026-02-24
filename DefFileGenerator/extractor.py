#!/usr/bin/env python3
import argparse
import csv
import json
import logging
import os
import re
import sys
import math

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

from DefFileGenerator.def_gen import Generator

class Extractor:
    COLUMN_MAPPING = {
        'RegisterType': ['register type', 'reg type', 'modbus type', 'registertype'],
        'Address': ['address', 'addr', 'offset', 'register', 'reg'],
        'Name': ['name', 'description', 'parameter', 'variable', 'signal'],
        'Type': ['data type', 'datatype', 'type', 'format'],
        'Unit': ['unit', 'units'],
        'Factor': ['scale', 'factor', 'multiplier', 'ratio'],
        'ScaleFactor': ['scalefactor'],
        'Tag': ['tag'],
        'Action': ['action', 'access']
    }

    TYPE_MAPPING = {
        'uint16': 'U16',
        'int16': 'I16',
        'uint32': 'U32',
        'int32': 'I32',
        'uint64': 'U64',
        'int64': 'I64',
        'float': 'F32',
        'f32': 'F32',
        'float32': 'F32',
        'double': 'F64',
        'f64': 'F64',
        'float64': 'F64',
        'string': 'STRING',
        'bits': 'BITS'
    }

    def __init__(self, mapping=None):
        self.mapping = mapping or {}
        self.generator = Generator()

    def normalize_type(self, t):
        if not t:
            return 'U16'
        t_str = str(t).lower().strip()

        for key, val in self.TYPE_MAPPING.items():
            if key in t_str:
                return val

        # Fallback to cleaning
        t_str = t_str.replace('unsigned ', 'u').replace('signed ', 'i').replace(' ', '')
        match = re.match(r'^(u|i|uint|int)(\d+)$', t_str)
        if match:
            raw_prefix = match.group(1).lower()
            prefix = 'U' if raw_prefix.startswith('u') else 'I'
            bits = match.group(2)
            return f"{prefix}{bits}"

        # Final cleanup
        t_str = re.sub(r'[^a-z0-9_]+', '', t_str)
        return t_str.upper() if t_str else 'U16'

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        tables = []
        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                content = f.read(2048)
                f.seek(0)
                # Manual delimiter check
                delimiter = ','
                for d in [';', '\t', ',']:
                    if d in content:
                        delimiter = d
                        break

                reader = csv.DictReader(f, delimiter=delimiter)
                rows = list(reader)
                if rows:
                    tables.append(rows)
        except Exception as e:
            logging.error(f"Error extracting from CSV: {e}")
        return tables

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL and not HAS_PANDAS:
            logging.error("openpyxl or pandas is required for Excel extraction.")
            return []
        logging.info(f"Extracting from Excel: {filepath}")
        tables = []
        try:
            if HAS_PANDAS:
                if sheet_name:
                    df = pd.read_excel(filepath, sheet_name=sheet_name)
                    tables.append(df.to_dict(orient='records'))
                else:
                    excel_file = pd.ExcelFile(filepath)
                    for sheet in excel_file.sheet_names:
                        df = excel_file.parse(sheet)
                        tables.append(df.to_dict(orient='records'))
            else:
                wb = openpyxl.load_workbook(filepath, data_only=True)
                sheets = [sheet_name] if sheet_name else wb.sheetnames
                for name in sheets:
                    ws = wb[name]
                    rows = list(ws.rows)
                    if not rows: continue
                    headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]
                    sheet_data = []
                    for row in rows[1:]:
                        row_data = {}
                        for i, cell in enumerate(row):
                            if i < len(headers):
                                row_data[headers[i]] = cell.value
                        sheet_data.append(row_data)
                    tables.append(sheet_data)
        except Exception as e:
            logging.error(f"Error extracting from Excel: {e}")
        return tables

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        logging.info(f"Extracting from PDF: {filepath}")
        tables = []
        try:
            with pdfplumber.open(filepath) as pdf:
                if pages is None:
                    target_pages = pdf.pages
                else:
                    target_pages = [pdf.pages[i-1] for i in pages if i <= len(pdf.pages)]

                for page in target_pages:
                    # extract_tables() returns all tables on the page
                    page_tables = page.extract_tables()
                    for table in page_tables:
                        if not table or len(table) < 2:
                            continue
                        headers = [str(c).replace('\n', ' ').strip() if c else "" for c in table[0]]
                        table_data = []
                        for row in table[1:]:
                            row_data = {}
                            for i, cell in enumerate(row):
                                if i < len(headers):
                                    val = str(cell).replace('\n', ' ').strip() if cell else ""
                                    row_data[headers[i]] = val
                            table_data.append(row_data)
                        tables.append(table_data)
        except Exception as e:
            logging.error(f"Error extracting from PDF: {e}")
        return tables

    def map_and_clean(self, tables):
        """Processes a list of tables and maps columns to standard names."""
        if not tables:
            return []

        # Handle single table input for backward compatibility
        if isinstance(tables, list) and tables and isinstance(tables[0], dict):
            tables = [tables]

        all_mapped_data = []

        for table in tables:
            if not table:
                continue

            # Identify mapping for this table
            first_row = table[0]
            col_map = {}
            used_src_cols = set()

            # 1. Explicit mapping from constructor
            for target, source in self.mapping.items():
                if source in first_row:
                    col_map[target] = source
                    used_src_cols.add(source)

            # 2. Heuristic mapping
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Factor', 'ScaleFactor', 'Tag']
            for target in detection_order:
                if target in col_map: continue
                for src_col in first_row.keys():
                    if src_col in used_src_cols: continue
                    col_lower = str(src_col).lower()
                    if any(pattern in col_lower for pattern in self.COLUMN_MAPPING.get(target, [])):
                        col_map[target] = src_col
                        used_src_cols.add(src_col)
                        break

            if 'Address' not in col_map and 'Name' not in col_map:
                continue

            for row in table:
                new_row = {}
                for target, source in col_map.items():
                    new_row[target] = row[source]

                # Basic cleaning
                if 'Name' not in new_row or not str(new_row.get('Name')).strip():
                    if 'Address' in new_row and new_row['Address']:
                        new_row['Name'] = f"Register {new_row['Address']}"
                    else:
                        continue

                # Normalize Address
                if 'Address' in new_row and new_row['Address']:
                    addr = str(new_row['Address']).strip()
                    # Remove commas
                    addr = addr.replace(',', '')
                    if '_' in addr:
                        parts = addr.split('_')
                        norm_parts = [self.generator.normalize_address_val(p) for p in parts]
                        new_row['Address'] = '_'.join(norm_parts)
                    else:
                        new_row['Address'] = self.generator.normalize_address_val(addr)

                # Normalize Type
                new_row['Type'] = self.normalize_type(new_row.get('Type', 'U16'))

                # Normalize Action
                if 'Action' in new_row:
                    new_row['Action'] = self.generator.normalize_action(new_row['Action'])

                all_mapped_data.append(new_row)

        return all_mapped_data

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

    parser = argparse.ArgumentParser(description='Extract register information from PDF/Excel/CSV files.')
    parser.add_argument('input_file', help='Path to the source PDF, Excel, or CSV file.')
    parser.add_argument('-o', '--output', help='Path to the output CSV file.')
    parser.add_argument('--mapping', help='JSON file containing column mapping.')
    parser.add_argument('--sheet', help='Excel sheet name to extract from.')
    parser.add_argument('--pages', help='PDF pages to extract from (comma separated, e.g. 1,2,5).')

    args = parser.parse_args()

    mapping = {}
    if args.mapping:
        try:
            with open(args.mapping, 'r') as f:
                mapping = json.load(f)
        except Exception as e:
            logging.error(f"Error loading mapping file: {e}")
            sys.exit(1)

    extractor = Extractor(mapping)

    ext = os.path.splitext(args.input_file)[1].lower()
    tables = []

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xls']:
        tables = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = None
        if args.pages:
            pages = [int(p.strip()) for p in args.pages.split(',')]
        tables = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv':
        tables = extractor.extract_from_csv(args.input_file)
    else:
        logging.error(f"Unsupported file extension: {ext}")
        sys.exit(1)

    if not tables:
        logging.error("No data extracted.")
        sys.exit(1)

    mapped_data = extractor.map_and_clean(tables)

    if not mapped_data:
        logging.error("No data remained after mapping and cleaning.")
        sys.exit(1)

    # Write to CSV
    output = args.output if args.output else sys.stdout

    fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']

    if isinstance(output, str):
        f = open(output, 'w', newline='', encoding='utf-8')
    else:
        f = output

    writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
    writer.writeheader()
    writer.writerows(mapped_data)

    if isinstance(output, str):
        f.close()
        logging.info(f"Extracted data saved to {args.output}")

if __name__ == "__main__":
    main()
