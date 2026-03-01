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
    # Centralized column mapping keywords
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

    def __init__(self, mapping=None):
        self.mapping = mapping or {}
        self.generator = Generator()

    def normalize_type(self, t):
        return self.generator.normalize_type(t)

    def normalize_action(self, a):
        return self.generator.normalize_action(a)

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
        logging.info(f"Extracting from Excel: {filepath}")
        wb = openpyxl.load_workbook(filepath, data_only=True)

        all_tables = []
        sheets = [sheet_name] if sheet_name else wb.sheetnames

        for name in sheets:
            if name not in wb.sheetnames:
                continue
            ws = wb[name]
            rows = list(ws.rows)
            if not rows:
                continue

            headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]
            table_data = []
            for row in rows[1:]:
                row_data = {headers[i]: cell.value for i, cell in enumerate(row) if i < len(headers)}
                table_data.append(row_data)
            all_tables.append(table_data)
        return all_tables

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        logging.info(f"Extracting from PDF: {filepath}")
        data = []
        with pdfplumber.open(filepath) as pdf:
            if pages is None:
                target_pages = pdf.pages
            else:
                target_pages = []
                # Simple page selection logic
                if isinstance(pages, int):
                    target_pages = [pdf.pages[pages-1]]
                elif isinstance(pages, list):
                    target_pages = [pdf.pages[i-1] for i in pages]

            for page in target_pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:
                        continue

                    # Clean headers: remove newlines
                    headers = [str(c).replace('\n', ' ').strip() if c else "" for c in table[0]]
                    page_table = []
                    for row in table[1:]:
                        row_data = {headers[i]: str(cell).replace('\n', ' ').strip() if cell else ""
                                   for i, cell in enumerate(row) if i < len(headers)}
                        page_table.append(row_data)
                    data.append(page_table)
        return data

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        # Detect delimiter
        with open(filepath, 'r', encoding='utf-8-sig') as f:
            content = f.read(2048)
            delimiter = ','
            for d in [';', '\t', ',']:
                if d in content:
                    delimiter = d
                    break
            f.seek(0)
            reader = csv.DictReader(f, delimiter=delimiter)
            return [list(reader)]

    def map_and_clean(self, raw_data):
        # Support both single table (list of dicts) and list of tables
        if raw_data and isinstance(raw_data[0], dict):
            tables = [raw_data]
        else:
            tables = raw_data

        all_mapped_data = []

        for table in tables:
            if not table:
                continue

            mapped_table = []
            first_row = table[0]
            standard_cols_mapping = {}
            used_src_cols = set()

            # 1. Explicit mapping
            for target, source in self.mapping.items():
                if source in first_row:
                    standard_cols_mapping[target] = source
                    used_src_cols.add(source)

            # 2. Heuristic mapping based on priority
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag', 'Factor', 'ScaleFactor']
            for target in detection_order:
                if target in standard_cols_mapping:
                    continue
                for src_col in first_row.keys():
                    if src_col in used_src_cols:
                        continue
                    src_lower = str(src_col).lower()
                    if any(kw in src_lower for kw in self.COLUMN_MAPPING.get(target, [])):
                        standard_cols_mapping[target] = src_col
                        used_src_cols.add(src_col)
                        break

            for row in table:
                new_row = {}
                for target, source in standard_cols_mapping.items():
                    val = row.get(source)
                    if val is not None:
                        new_row[target] = str(val).strip()

                if not new_row.get('Name') and not new_row.get('Address'):
                    continue

                # Clean Address
                if new_row.get('Address'):
                    addr = new_row['Address']
                    if '_' in addr:
                        parts = addr.split('_')
                        new_row['Address'] = '_'.join([self.generator.normalize_address_val(p) for p in parts])
                    else:
                        new_row['Address'] = self.generator.normalize_address_val(addr)

                # Clean Type & Action
                new_row['Type'] = self.normalize_type(new_row.get('Type'))
                if 'Action' in new_row:
                    new_row['Action'] = self.normalize_action(new_row['Action'])

                if 'RegisterType' not in new_row:
                    new_row['RegisterType'] = 'Holding Register'

                mapped_table.append(new_row)
            all_mapped_data.extend(mapped_table)

        return all_mapped_data

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

    parser = argparse.ArgumentParser(description='Extract register information from PDF/Excel files.')
    parser.add_argument('input_file', help='Path to the source PDF or Excel file.')
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
    raw_data = []

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xls']:
        raw_data = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.csv':
        raw_data = extractor.extract_from_csv(args.input_file)
    elif ext == '.pdf':
        pages = None
        if args.pages:
            pages = [int(p.strip()) for p in args.pages.split(',')]
        raw_data = extractor.extract_from_pdf(args.input_file, pages)
    else:
        logging.error(f"Unsupported file extension: {ext}")
        sys.exit(1)

    if not raw_data:
        logging.error("No data extracted.")
        sys.exit(1)

    mapped_data = extractor.map_and_clean(raw_data)

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
