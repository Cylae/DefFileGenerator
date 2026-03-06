#!/usr/bin/env python3
import argparse
import csv
import json
import logging
import os
import re
import sys

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

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
        self.generator = Generator()

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
        logging.info(f"Extracting from Excel: {filepath}")
        wb = openpyxl.load_workbook(filepath, data_only=True)

        sheets_to_process = []
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                logging.error(f"Sheet '{sheet_name}' not found in {filepath}")
                return []
            sheets_to_process = [wb[sheet_name]]
        else:
            sheets_to_process = wb.worksheets

        all_tables = []
        for ws in sheets_to_process:
            rows = list(ws.rows)
            if not rows:
                continue

            headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]
            table_data = []
            for row in rows[1:]:
                row_data = {}
                for i, cell in enumerate(row):
                    if i < len(headers):
                        row_data[headers[i]] = cell.value
                table_data.append(row_data)
            if table_data:
                all_tables.append(table_data)
        return all_tables

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        logging.info(f"Extracting from PDF: {filepath}")
        all_tables = []
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
                    table_data = []
                    for row in table[1:]:
                        row_data = {}
                        for i, cell in enumerate(row):
                            if i < len(headers):
                                val = str(cell).replace('\n', ' ').strip() if cell else ""
                                row_data[headers[i]] = val
                        table_data.append(row_data)
                    if table_data:
                        all_tables.append(table_data)
        return all_tables

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        all_tables = []
        for delimiter in [',', ';', '\t']:
            try:
                with open(filepath, 'r', encoding='utf-8-sig') as f:
                    content = f.read(2048)
                    f.seek(0)
                    # Manual delimiter check
                    if delimiter not in content:
                        continue

                    reader = csv.DictReader(f, delimiter=delimiter)
                    rows = list(reader)
                    if rows and len(reader.fieldnames) > 1:
                        all_tables.append(rows)
                        return all_tables
            except Exception:
                continue
        return []

    def extract_from_xml(self, filepath):
        if not HAS_PANDAS:
            logging.error("pandas is required for XML extraction.")
            return []
        logging.info(f"Extracting from XML: {filepath}")
        try:
            df = pd.read_xml(filepath)
            return [df.to_dict(orient='records')]
        except Exception as e:
            logging.debug(f"Default XML parser failed, trying etree: {e}")
            try:
                df = pd.read_xml(filepath, parser='etree')
                return [df.to_dict(orient='records')]
            except Exception as e2:
                logging.error(f"Error loading XML file: {e2}")
                return []

    def map_and_clean(self, tables):
        # Handle both single table (list of dicts) or list of tables
        if tables and isinstance(tables[0], dict):
            tables = [tables]

        all_mapped_data = []

        for raw_data in tables:
            if not raw_data:
                continue

            # Identify standard columns for this table
            first_row = raw_data[0]
            standard_cols_mapping = {}
            used_src_cols = set()

            # 1. Explicitly mapped columns from config
            for target, source in self.mapping.items():
                if source in first_row:
                    standard_cols_mapping[target] = source
                    used_src_cols.add(source)

            # 2. Fuzzy match for standard columns
            # Priority order: RegisterType > Address > Name > Type > Unit > Action > Tag
            priority_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag', 'Factor', 'ScaleFactor']
            for target in priority_order:
                if target in standard_cols_mapping:
                    continue
                for k in first_row.keys():
                    if k in used_src_cols:
                        continue
                    k_lower = str(k).lower()
                    for pattern in self.COLUMN_MAPPING.get(target, []):
                        if pattern in k_lower:
                            standard_cols_mapping[target] = k
                            used_src_cols.add(k)
                            break
                    if target in standard_cols_mapping:
                        break

            for row in raw_data:
                new_row = {}
                for target, source in standard_cols_mapping.items():
                    val = row.get(source)
                    if val is not None:
                        new_row[target] = val

                # Special handling for Factor (fractions like 1/10)
                if 'Factor' in new_row and new_row['Factor']:
                    factor_val = str(new_row['Factor'])
                    if '/' in factor_val:
                        try:
                            parts = factor_val.split('/')
                            new_row['Factor'] = str(float(parts[0]) / float(parts[1]))
                        except (ValueError, ZeroDivisionError):
                            pass

                # Clean Address
                if 'Address' in new_row and new_row['Address']:
                    addr = str(new_row['Address']).strip()
                    if '_' in addr:
                        parts = addr.split('_')
                        norm_parts = [self.generator.normalize_address_val(p) for p in parts]
                        new_row['Address'] = '_'.join(norm_parts)
                    else:
                        new_row['Address'] = self.generator.normalize_address_val(addr)

                # Delegate type and action normalization to Generator when processing rows,
                # but we can do a first pass here for the simplified CSV if we want.
                # However, memories say: "To avoid redundancy, it leaves data type normalization to the Generator.process_rows method."

                if 'Name' not in new_row or not new_row['Name']:
                    # If we have an address but no name, we can still use it.
                    # If both missing, skip.
                    if 'Address' not in new_row or not new_row['Address']:
                        continue

                if 'RegisterType' not in new_row:
                    new_row['RegisterType'] = 'Holding Register'

                all_mapped_data.append(new_row)

        return all_mapped_data

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

    parser = argparse.ArgumentParser(description='Extract register information from PDF/Excel/CSV/XML files.')
    parser.add_argument('input_file', help='Path to the source file.')
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
    elif ext == '.xml':
        tables = extractor.extract_from_xml(args.input_file)
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
