#!/usr/bin/env python3
import argparse
import csv
import json
import logging
import os
import re
import sys
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

    TYPE_PATTERN = re.compile(r'^(u|i|uint|int)(\d+)$', re.IGNORECASE)

    def __init__(self, mapping=None):
        self.mapping = mapping or {}
        self.generator = Generator()

    def normalize_type(self, t):
        if not t:
            return 'U16'
        t_str = str(t).lower().strip()

        # Use the generator's robust normalization logic
        return self.generator.normalize_type(t_str)

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
        logging.info(f"Extracting from Excel: {filepath}")
        wb = openpyxl.load_workbook(filepath, data_only=True)

        all_data = []
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                logging.error(f"Sheet '{sheet_name}' not found.")
                return []
            sheets = [wb[sheet_name]]
        else:
            sheets = wb.worksheets

        for ws in sheets:
            rows = list(ws.rows)
            if not rows:
                continue

            headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]
            for row in rows[1:]:
                row_data = {}
                for i, cell in enumerate(row):
                    if i < len(headers):
                        row_data[headers[i]] = cell.value
                all_data.append(row_data)
        return all_data

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        logging.info(f"Extracting from PDF: {filepath}")
        all_data = []
        with pdfplumber.open(filepath) as pdf:
            if pages is None:
                target_pages = pdf.pages
            else:
                target_pages = [pdf.pages[i-1] for i in pages if 0 < i <= len(pdf.pages)]

            for page in target_pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:
                        continue

                    headers = [str(c).replace('\n', ' ').strip() if c else "" for c in table[0]]
                    for row in table[1:]:
                        row_data = {}
                        for i, cell in enumerate(row):
                            if i < len(headers):
                                val = str(cell).replace('\n', ' ').strip() if cell else ""
                                row_data[headers[i]] = val
                        all_data.append(row_data)
        return all_data

    def map_and_clean(self, raw_data):
        if not raw_data:
            return []

        # If raw_data is a list of tables, flatten it for now or process each
        # The existing tests expect flattened data for some calls, or just one table
        if raw_data and isinstance(raw_data[0], list):
            # It's a list of tables
            tables = raw_data
        else:
            # It's a single table (list of dicts)
            tables = [raw_data]

        all_mapped_data = []

        for table in tables:
            if not table:
                continue

            first_row = table[0]
            standard_cols_mapping = {}
            used_src_cols = set()

            # Explicitly mapped columns
            for target, source in self.mapping.items():
                if source in first_row:
                    standard_cols_mapping[target] = source
                    used_src_cols.add(source)

            # Fuzzy match
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Factor', 'ScaleFactor', 'Tag', 'Action']
            for target in detection_order:
                if target in standard_cols_mapping:
                    continue
                for k in first_row.keys():
                    if k in used_src_cols:
                        continue
                    k_lower = k.lower()
                    for pattern in self.COLUMN_MAPPING.get(target, []):
                        if pattern in k_lower:
                            standard_cols_mapping[target] = k
                            used_src_cols.add(k)
                            break
                    if target in standard_cols_mapping:
                        break

            for row in table:
                new_row = {}
                for target, source in standard_cols_mapping.items():
                    val = row.get(source)
                    new_row[target] = val

                # Extra columns
                for k, v in row.items():
                    if k not in used_src_cols and k not in new_row:
                        new_row[k] = v

                if not new_row.get('Name'):
                    continue

                # Normalize Address
                addr = str(new_row.get('Address', '')).strip()
                if addr:
                    if '_' in addr:
                        parts = addr.split('_')
                        new_row['Address'] = '_'.join([self.generator.normalize_address_val(p) for p in parts])
                    else:
                        new_row['Address'] = self.generator.normalize_address_val(addr)

                # Normalize Type
                new_row['Type'] = self.normalize_type(new_row.get('Type', 'U16'))

                # Normalize Factor (fractions like 1/10)
                factor = str(new_row.get('Factor', '1'))
                if '/' in factor:
                    try:
                        parts = factor.split('/')
                        new_row['Factor'] = str(float(parts[0]) / float(parts[1]))
                    except (ValueError, ZeroDivisionError):
                        new_row['Factor'] = '1'

                if 'RegisterType' not in new_row:
                    new_row['RegisterType'] = 'Holding Register'

                all_mapped_data.append(new_row)

        return all_mapped_data

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(description='Extract registers from PDF/Excel.')
    parser.add_argument('input_file', help='Source file.')
    parser.add_argument('-o', '--output', help='Output CSV.')
    parser.add_argument('--mapping', help='Mapping JSON.')
    parser.add_argument('--sheet', help='Excel sheet.')
    parser.add_argument('--pages', help='PDF pages.')

    args = parser.parse_args()
    mapping = {}
    if args.mapping:
        with open(args.mapping, 'r') as f:
            mapping = json.load(f)

    extractor = Extractor(mapping)
    ext = os.path.splitext(args.input_file)[1].lower()

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xls']:
        raw_data = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = [int(p.strip()) for p in args.pages.split(',')] if args.pages else None
        raw_data = extractor.extract_from_pdf(args.input_file, pages)
    else:
        logging.error(f"Unsupported extension: {ext}")
        sys.exit(1)

    mapped_data = extractor.map_and_clean(raw_data)
    if not mapped_data:
        logging.error("No data extracted.")
        sys.exit(1)

    fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
    output = args.output if args.output else sys.stdout
    if isinstance(output, str):
        f = open(output, 'w', newline='', encoding='utf-8')
    else:
        f = output

    writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
    writer.writeheader()
    writer.writerows(mapped_data)
    if isinstance(output, str):
        f.close()

if __name__ == "__main__":
    main()
