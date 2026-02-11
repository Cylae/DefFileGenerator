#!/usr/bin/env python3
import argparse
import csv
import json
import logging
import os
import re
import sys
import openpyxl
import pdfplumber
from DefFileGenerator.def_gen import Generator

class Extractor:
    def __init__(self, mapping=None):
        self.mapping = mapping or {}
        self.generator = Generator()
        self.column_patterns = {
            'RegisterType': ['register type', 'reg type', 'modbus type', 'registertype'],
            'Address': ['address', 'addr', 'offset', 'register', 'reg'],
            'Name': ['name', 'description', 'parameter', 'variable', 'signal'],
            'Type': ['data type', 'datatype', 'type', 'format'],
            'Unit': ['unit', 'units'],
            'Factor': ['scale', 'factor', 'multiplier', 'ratio'],
            'Action': ['action', 'access']
        }

    def normalize_type(self, t):
        """Uses centralized Generator logic for type normalization."""
        return self.generator.normalize_type(t)

    def extract_from_excel(self, filepath, sheet_name=None):
        logging.info(f"Extracting from Excel: {filepath}")
        wb = openpyxl.load_workbook(filepath, data_only=True)
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                logging.error(f"Sheet '{sheet_name}' not found in {filepath}")
                return []
            ws = wb[sheet_name]
        else:
            ws = wb.active

        data = []
        rows = list(ws.rows)
        if not rows:
            return []

        headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]

        for row_idx, row in enumerate(rows[1:], start=2):
            row_data = {}
            for i, cell in enumerate(row):
                if i < len(headers):
                    row_data[headers[i]] = cell.value
            data.append(row_data)
        return data

    def extract_from_pdf(self, filepath, pages=None):
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

                    for row in table[1:]:
                        row_data = {}
                        for i, cell in enumerate(row):
                            if i < len(headers):
                                val = str(cell).replace('\n', ' ').strip() if cell else ""
                                row_data[headers[i]] = val
                        data.append(row_data)
        return data

    def map_and_clean(self, raw_data):
        mapped_data = []
        if not raw_data:
            return []

        # Identify standard columns once
        first_row = raw_data[0]
        standard_cols_mapping = {}
        assigned_src_cols = set()

        # 1. Explicitly mapped columns from config
        for target, source in self.mapping.items():
            if source in first_row:
                standard_cols_mapping[target] = source
                assigned_src_cols.add(source)

        # 2. Heuristic match for standard columns if not explicitly mapped
        # Priority order to avoid misidentification
        targets = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Factor', 'Action', 'Tag']
        for target in targets:
            if target in standard_cols_mapping:
                continue

            patterns = self.column_patterns.get(target, [target.lower()])
            for src_col in first_row.keys():
                if src_col in assigned_src_cols:
                    continue

                src_col_lower = str(src_col).lower()
                if any(p in src_col_lower for p in patterns):
                    standard_cols_mapping[target] = src_col
                    assigned_src_cols.add(src_col)
                    break

        for row in raw_data:
            new_row = {}
            # Apply identified mappings
            for target, source in standard_cols_mapping.items():
                val = row.get(source)
                if val is not None:
                    new_row[target] = val

            # Fill in any remaining columns
            for k, v in row.items():
                if k not in assigned_src_cols and k not in new_row:
                    new_row[k] = v

            # Normalization
            # 1. Address
            if 'Address' in new_row and new_row['Address'] is not None:
                addr = str(new_row['Address']).strip()
                if '_' in addr:
                    parts = addr.split('_')
                    norm_parts = [self.generator.normalize_address_val(p) for p in parts]
                    new_row['Address'] = '_'.join(norm_parts)
                else:
                    new_row['Address'] = self.generator.normalize_address_val(addr)

            # 2. Type
            if 'Type' in new_row:
                new_row['Type'] = self.normalize_type(new_row['Type'])

            # 3. Action
            if 'Action' in new_row:
                new_row['Action'] = self.generator.normalize_action(new_row['Action'])

            # 4. Factor (handle fractions like 1/10)
            if 'Factor' in new_row and new_row['Factor'] is not None:
                factor_str = str(new_row['Factor']).strip()
                if '/' in factor_str:
                    try:
                        p1, p2 = factor_str.split('/')
                        new_row['Factor'] = str(float(p1) / float(p2))
                    except (ValueError, ZeroDivisionError):
                        pass

            # Skip rows without Name and Address
            if not new_row.get('Name') and not new_row.get('Address'):
                continue

            # Ensure RegisterType has a default if missing
            if 'RegisterType' not in new_row or not new_row['RegisterType']:
                new_row['RegisterType'] = 'Holding Register'

            mapped_data.append(new_row)

        return mapped_data

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

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw_data = extractor.extract_from_excel(args.input_file, args.sheet)
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
