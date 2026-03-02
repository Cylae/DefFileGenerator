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
        'Scale': ['scale', 'factor', 'multiplier', 'ratio'],
        'Action': ['action', 'access'],
        'Tag': ['tag']
    }

    def __init__(self, mapping=None):
        self.mapping = mapping or {}

    def extract_from_excel(self, filepath, sheet_name=None):
        logging.info(f"Extracting from Excel: {filepath}")
        all_tables = []

        if HAS_PANDAS:
            try:
                if sheet_name:
                    df = pd.read_excel(filepath, sheet_name=sheet_name)
                    all_tables.append(df.to_dict(orient='records'))
                else:
                    excel_file = pd.ExcelFile(filepath)
                    for name in excel_file.sheet_names:
                        df = excel_file.parse(name)
                        all_tables.append(df.to_dict(orient='records'))
                return all_tables
            except Exception as e:
                logging.warning(f"Pandas Excel extraction failed: {e}. Trying openpyxl...")

        if not HAS_OPENPYXL:
            logging.error("Neither pandas nor openpyxl is available for Excel extraction.")
            return []

        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            sheets_to_process = [sheet_name] if sheet_name else wb.sheetnames

            for name in sheets_to_process:
                if name not in wb.sheetnames:
                    logging.warning(f"Sheet '{name}' not found in {filepath}")
                    continue
                ws = wb[name]
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
                    all_tables.append(data)
            return all_tables
        except Exception as e:
            logging.error(f"Error extracting from Excel with openpyxl: {e}")
            return []

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
                for p in pages:
                    if 1 <= p <= len(pdf.pages):
                        target_pages.append(pdf.pages[p-1])

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
                        all_tables.append(data)
        return all_tables

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        for delimiter in [',', ';', '\t']:
            try:
                with open(filepath, 'r', encoding='utf-8-sig') as f:
                    sample = f.read(2048)
                    f.seek(0)
                    dialect = csv.Sniffer().sniff(sample, delimiters=delimiter)
                    reader = csv.DictReader(f, dialect=dialect)
                    rows = list(reader)
                    if rows and len(reader.fieldnames) > 1:
                        return [rows]
            except Exception:
                continue
        # Fallback
        for delimiter in [',', ';', '\t']:
            try:
                with open(filepath, 'r', encoding='utf-8-sig') as f:
                    reader = csv.DictReader(f, delimiter=delimiter)
                    rows = list(reader)
                    if rows and len(reader.fieldnames) > 1:
                        return [rows]
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
        """
        Processes raw extracted data (list of tables) into a flat list of mapped registers.
        """
        # Support both single table or list of tables
        if tables and isinstance(tables[0], dict):
            tables = [tables]

        mapped_data = []
        generator = Generator()

        for table in tables:
            if not table:
                continue

            # Identify columns for this table
            col_map = {}
            assigned_src_cols = set()
            first_row = table[0]

            # 1. Explicit mapping from config
            for target, source in self.mapping.items():
                if source in first_row:
                    col_map[target] = source
                    assigned_src_cols.add(source)

            # 2. Priority-based fuzzy matching
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Scale', 'Action', 'Tag']
            for target in detection_order:
                if target in col_map:
                    continue

                patterns = self.COLUMN_MAPPING.get(target, [target.lower()])
                for src_col in first_row.keys():
                    if src_col in assigned_src_cols:
                        continue

                    src_col_lower = str(src_col).lower()
                    if any(p in src_col_lower for p in patterns):
                        col_map[target] = src_col
                        assigned_src_cols.add(src_col)
                        break

            if 'Address' not in col_map and 'Name' not in col_map:
                logging.debug("Skipping table as neither Address nor Name columns found.")
                continue

            for row in table:
                # Basic validation: must have some content
                if not any(v for v in row.values() if v):
                    continue

                new_row = {}

                # Extract using mapping
                for target, src_col in col_map.items():
                    val = row.get(src_col)
                    if val is not None:
                        new_row[target] = val

                # Clean Address
                if 'Address' in new_row and new_row['Address']:
                    addr = str(new_row['Address']).strip()
                    if '_' in addr:
                        parts = addr.split('_')
                        norm_parts = [generator.normalize_address_val(p) for p in parts]
                        new_row['Address'] = '_'.join(norm_parts)
                    else:
                        new_row['Address'] = generator.normalize_address_val(addr)

                # Map Scale to Factor if Factor not already present
                if 'Scale' in new_row and 'Factor' not in new_row:
                    new_row['Factor'] = new_row.pop('Scale')

                # Clean Type
                # Note: Generator.process_rows also performs normalization, so we leave it raw here

                # Skip rows without a name or address
                if (not new_row.get('Name')) and (not new_row.get('Address')):
                    continue

                # If name is missing but address is present, generate dummy name
                if not new_row.get('Name') and new_row.get('Address'):
                    new_row['Name'] = f"Register {new_row['Address']}"

                # Ensure some defaults for Generator
                if 'RegisterType' not in new_row:
                    new_row['RegisterType'] = 'Holding Register'

                mapped_data.append(new_row)

        return mapped_data

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
    raw_data = []

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xls']:
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
