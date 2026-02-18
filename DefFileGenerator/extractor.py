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
        'Action': ['action', 'access'],
        'Tag': ['tag'],
        'Offset': ['offset_val'], # To avoid confusion with Address 'offset'
        'ScaleFactor': ['scalefactor', 'scale factor']
    }

    def __init__(self, mapping=None):
        self.mapping = mapping or {}
        self.generator = Generator()

    def normalize_type(self, t):
        """Delegates type normalization to the Generator class."""
        return self.generator.normalize_type(t)

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
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
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        logging.info(f"Extracting from PDF: {filepath}")
        data = []
        try:
            with pdfplumber.open(filepath) as pdf:
                if pages is None:
                    target_pages = pdf.pages
                else:
                    target_pages = []
                    for p in pages:
                        if 0 < p <= len(pdf.pages):
                            target_pages.append(pdf.pages[p-1])

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
        except Exception as e:
            logging.error(f"Error extracting from PDF: {e}")
        return data

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        data = []
        try:
            with open(filepath, mode='r', encoding='utf-8-sig') as f:
                content = f.read(4096)
                f.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(content, delimiters=";,")
                except csv.Error:
                    dialect = csv.excel
                    dialect.delimiter = ',' if ',' in content else ';'

                reader = csv.DictReader(f, dialect=dialect)
                for row in reader:
                    data.append(row)
        except Exception as e:
            logging.error(f"Error reading CSV {filepath}: {e}")
        return data

    def extract_from_xml(self, filepath):
        if not HAS_PANDAS:
            logging.error("pandas and lxml are required for XML extraction.")
            return []
        logging.info(f"Extracting from XML: {filepath}")
        try:
            # Try with default parser (usually lxml if installed)
            df = pd.read_xml(filepath)
            return df.to_dict(orient='records')
        except Exception as e:
            logging.error(f"Error reading XML {filepath}: {e}")
        return []

    def map_and_clean(self, raw_data):
        mapped_data = []
        if not raw_data:
            return []

        # Identify standard columns once to avoid repeated fuzzy matching
        first_row = raw_data[0]
        standard_cols_mapping = {}
        used_src_cols = set()

        # 1. Explicitly mapped columns from config
        for target, source in self.mapping.items():
            if source in first_row:
                standard_cols_mapping[target] = source
                used_src_cols.add(source)

        # 2. Fuzzy match for standard columns if not explicitly mapped
        detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Tag', 'Factor', 'Offset', 'Action', 'ScaleFactor']

        for target in detection_order:
            if target in standard_cols_mapping:
                continue

            target_patterns = self.COLUMN_MAPPING.get(target, [target.lower()])

            for k in first_row.keys():
                if k in used_src_cols:
                    continue
                k_lower = str(k).lower()

                # Check for exact match or pattern in key
                if any(p == k_lower or p in k_lower for p in target_patterns):
                    standard_cols_mapping[target] = k
                    used_src_cols.add(k)
                    break

        for row in raw_data:
            new_row = {}
            for target, source in standard_cols_mapping.items():
                if source in row:
                    new_row[target] = row[source]

            for k, v in row.items():
                if k not in used_src_cols and k not in new_row:
                    new_row[k] = v

            if 'Address' in new_row and new_row['Address'] is not None:
                addr = str(new_row['Address']).strip().replace(',', '')
                if '_' in addr:
                    parts = addr.split('_')
                    norm_parts = [self.generator.normalize_address_val(p) for p in parts]
                    new_row['Address'] = '_'.join(norm_parts)
                else:
                    new_row['Address'] = self.generator.normalize_address_val(addr)

            if 'Type' in new_row:
                new_row['Type'] = self.generator.normalize_type(new_row['Type'])

            if ('Name' not in new_row or not new_row['Name']) and ('Address' not in new_row or not new_row['Address']):
                continue

            if 'RegisterType' not in new_row:
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
    elif ext == '.csv':
        raw_data = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml':
        raw_data = extractor.extract_from_xml(args.input_file)
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
