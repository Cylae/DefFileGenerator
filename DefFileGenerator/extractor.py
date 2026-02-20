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

    def __init__(self, mapping=None):
        self.mapping = mapping or {}

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
        logging.info(f"Extracting from Excel: {filepath}")
        wb = openpyxl.load_workbook(filepath, data_only=True)

        tables = []
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                logging.error(f"Sheet '{sheet_name}' not found in {filepath}")
                return []
            sheets = [wb[sheet_name]]
        else:
            sheets = wb.worksheets

        for ws in sheets:
            data = []
            rows = list(ws.rows)
            if not rows:
                continue
            headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]
            for row in rows[1:]:
                row_data = {}
                for i, cell in enumerate(row):
                    if i < len(headers):
                        row_data[headers[i]] = cell.value
                data.append(row_data)
            if data:
                tables.append(data)
        return tables

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        logging.info(f"Extracting from PDF: {filepath}")
        tables_data = []
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
                    data = []
                    for row in table[1:]:
                        row_data = {}
                        for i, cell in enumerate(row):
                            if i < len(headers):
                                val = str(cell).replace('\n', ' ').strip() if cell else ""
                                row_data[headers[i]] = val
                        data.append(row_data)
                    if data:
                        tables_data.append(data)
        return tables_data

    def map_and_clean(self, tables):
        if not tables:
            return []

        # Handle backward compatibility if single table passed
        if isinstance(tables, list) and tables and isinstance(tables[0], dict):
            tables = [tables]

        all_mapped_data = []
        generator = Generator()

        for table in tables:
            if not table:
                continue

            first_row = table[0]
            col_map = {}
            used_src_cols = set()

            # Explicitly mapped columns from config
            for target, source in self.mapping.items():
                if source in first_row:
                    col_map[target] = source
                    used_src_cols.add(source)

            # Fuzzy match for standard columns
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Tag', 'Factor', 'ScaleFactor', 'Action']
            for target in detection_order:
                if target in col_map:
                    continue
                for k in first_row.keys():
                    if k in used_src_cols:
                        continue
                    k_lower = str(k).lower()
                    for pattern in self.COLUMN_MAPPING.get(target, []):
                        if pattern in k_lower:
                            col_map[target] = k
                            used_src_cols.add(k)
                            break
                    if target in col_map:
                        break

            if 'Address' not in col_map and 'Name' not in col_map:
                continue

            for row in table:
                new_row = {}
                for target, source in col_map.items():
                    new_row[target] = row[source]

                # Copy other columns
                for k, v in row.items():
                    if k not in used_src_cols:
                        new_row[k] = v

                # Use Generator for final normalization where appropriate
                if 'Name' not in new_row or not new_row['Name']:
                    # Use Address as name if name missing
                    if 'Address' in new_row and new_row['Address']:
                        new_row['Name'] = f"Register {new_row['Address']}"
                    else:
                        continue

                # Delegate final cleaning to Generator (will happen in process_rows)
                # But we can do basic type normalization here to help
                if 'Type' in new_row:
                    new_row['Type'] = generator.normalize_type(new_row['Type'])

                all_mapped_data.append(new_row)

        return all_mapped_data

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(description='Extract register information from PDF/Excel files.')
    parser.add_argument('input_file', help='Path to the source PDF or Excel file.')
    parser.add_argument('-o', '--output', help='Path to the output CSV file.')
    parser.add_argument('--mapping', help='JSON file containing column mapping.')
    parser.add_argument('--sheet', help='Excel sheet name to extract from.')
    parser.add_argument('--pages', help='PDF pages to extract from (comma separated).')

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

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        tables = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = [int(p.strip()) for p in args.pages.split(',')] if args.pages else None
        tables = extractor.extract_from_pdf(args.input_file, pages)
    else:
        logging.error(f"Unsupported file extension: {ext}")
        sys.exit(1)

    mapped_data = extractor.map_and_clean(tables)
    if not mapped_data:
        logging.error("No data extracted.")
        sys.exit(1)

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
