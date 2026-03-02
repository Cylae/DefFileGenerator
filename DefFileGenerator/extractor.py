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
        'Action': ['action', 'access'],
        'Factor': ['scale', 'factor', 'multiplier', 'ratio'],
        'ScaleFactor': ['scalefactor'],
        'Tag': ['tag']
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

        tables = []
        if sheet_name:
            sheets = [sheet_name] if sheet_name in wb.sheetnames else []
            if not sheets:
                logging.error(f"Sheet '{sheet_name}' not found in {filepath}")
                return []
        else:
            sheets = wb.sheetnames

        for name in sheets:
            ws = wb[name]
            data = []
            rows = list(ws.rows)
            if not rows or len(rows) < 2:
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

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        tables = []
        for delimiter in [',', ';', '\t']:
            try:
                with open(filepath, 'r', encoding='utf-8-sig') as f:
                    reader = csv.DictReader(f, delimiter=delimiter)
                    data = list(reader)
                    if data and len(reader.fieldnames) > 1:
                        # Normalize headers
                        for row in data:
                            for k in list(row.keys()):
                                if k is not None:
                                    row[k.strip()] = row.pop(k)
                        tables.append(data)
                        break
            except Exception:
                continue
        return tables

    def extract_from_xml(self, filepath):
        if not HAS_PANDAS:
            logging.error("pandas and lxml/etree are required for XML extraction.")
            return []
        logging.info(f"Extracting from XML: {filepath}")
        try:
            df = pd.read_xml(filepath)
            return [df.to_dict(orient='records')]
        except Exception:
            try:
                df = pd.read_xml(filepath, parser='etree')
                return [df.to_dict(orient='records')]
            except Exception as e:
                logging.error(f"Error loading XML file: {e}")
                return []

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
                target_pages = []
                if isinstance(pages, int):
                    target_pages = [pdf.pages[pages-1]]
                elif isinstance(pages, list):
                    target_pages = [pdf.pages[i-1] for i in pages]

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

    def map_and_clean(self, raw_data):
        if not raw_data:
            return []

        # Handle both single table (list of dicts) or list of tables (list of list of dicts)
        if isinstance(raw_data[0], list):
            all_mapped = []
            for table in raw_data:
                all_mapped.extend(self._process_single_table(table))
            return all_mapped
        else:
            return self._process_single_table(raw_data)

    def _process_single_table(self, table):
        if not table:
            return []

        mapped_data = []
        # Identify standard columns for this table
        first_row = table[0]
        standard_cols_mapping = {}
        used_src_cols = set()

        # Explicitly mapped columns from config
        for target, source in self.mapping.items():
            if source in first_row:
                standard_cols_mapping[target] = source
                used_src_cols.add(source)

        # Priority-based fuzzy match
        detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Factor', 'ScaleFactor', 'Tag']
        for target in detection_order:
            if target in standard_cols_mapping:
                continue
            patterns = self.COLUMN_MAPPING.get(target, [target.lower()])
            for k in first_row.keys():
                if k in used_src_cols:
                    continue
                k_lower = str(k).lower()
                if any(p in k_lower for p in patterns):
                    standard_cols_mapping[target] = k
                    used_src_cols.add(k)
                    break

        for row in table:
            new_row = {}
            for target, source in standard_cols_mapping.items():
                new_row[target] = row.get(source)

            # Mandatory fields and cleaning
            addr_raw = new_row.get('Address')
            name_raw = new_row.get('Name')

            if addr_raw is None and name_raw is None:
                continue

            # Clean Address
            if addr_raw:
                addr_str = str(addr_raw).strip().replace(',', '')
                # Basic normalization here, detailed parsing in Generator.process_rows
                if '_' in addr_str:
                    parts = addr_str.split('_')
                    norm_parts = [self.generator.normalize_address_val(p) for p in parts]
                    new_row['Address'] = '_'.join(norm_parts)
                else:
                    new_row['Address'] = self.generator.normalize_address_val(addr_str)

            # Normalize Type
            if 'Type' in new_row:
                new_row['Type'] = self.normalize_type(new_row['Type'])
            else:
                new_row['Type'] = 'U16'

            # Normalize Action
            if 'Action' in new_row:
                new_row['Action'] = self.normalize_action(new_row['Action'])

            # RegisterType default
            if not new_row.get('RegisterType'):
                new_row['RegisterType'] = 'Holding Register'

            if not new_row.get('Name') and new_row.get('Address'):
                new_row['Name'] = f"Register {new_row['Address']}"

            if new_row.get('Name'):
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
