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
    def __init__(self, mapping=None):
        self.mapping = mapping or {}

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
        logging.info(f"Extracting from Excel: {filepath}")
        wb = openpyxl.load_workbook(filepath, data_only=True)

        all_tables = []
        target_sheets = [sheet_name] if sheet_name else wb.sheetnames

        for name in target_sheets:
            if name not in wb.sheetnames:
                logging.warning(f"Sheet '{name}' not found.")
                continue
            ws = wb[name]
            table_data = []
            rows = list(ws.rows)
            if not rows:
                continue

            headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]
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
                for p in pages:
                    if 1 <= p <= len(pdf.pages):
                        target_pages.append(pdf.pages[p-1])

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
        # Manual delimiter detection
        delimiters = [',', ';', '\t']
        best_dialect = None

        try:
            # Detect encoding (UTF-16 vs UTF-8)
            encoding = 'utf-8-sig'
            try:
                with open(filepath, 'rb') as f:
                    header = f.read(2)
                    if header in (b'\xff\xfe', b'\xfe\xff'):
                        encoding = 'utf-16'
            except Exception:
                pass

            with open(filepath, 'r', encoding=encoding) as f:
                sample = f.read(2048)
                if not sample:
                    return []
                f.seek(0)
                try:
                    best_dialect = csv.Sniffer().sniff(sample, delimiters=";,")
                except csv.Error:
                    # Fallback to manual check
                    max_cols = 0
                    for d in delimiters:
                        f.seek(0)
                        reader = csv.reader(f, delimiter=d)
                        try:
                            first_row = next(reader)
                            if len(first_row) > max_cols:
                                max_cols = len(first_row)
                                best_dialect = csv.excel
                                best_dialect.delimiter = d
                        except StopIteration:
                            continue

            if not best_dialect:
                return []

            with open(filepath, 'r', encoding=encoding) as f:
                reader = csv.DictReader(f, dialect=best_dialect)
                table_data = list(reader)
                if table_data:
                    all_tables.append(table_data)
        except Exception as e:
            logging.error(f"Error extracting from CSV: {e}")

        return all_tables

    def extract_from_xml(self, filepath):
        if not HAS_PANDAS:
            logging.error("pandas is required for XML extraction.")
            return []
        logging.info(f"Extracting from XML: {filepath}")
        all_tables = []
        try:
            df = pd.read_xml(filepath)
            table_data = df.to_dict(orient='records')
            if table_data:
                all_tables.append(table_data)
        except Exception as e:
            logging.debug(f"Default XML parser failed, trying etree: {e}")
            try:
                df = pd.read_xml(filepath, parser='etree')
                table_data = df.to_dict(orient='records')
                if table_data:
                    all_tables.append(table_data)
            except Exception as e2:
                logging.error(f"Error extracting from XML: {e2}")
        return all_tables

    # Static column mapping for priority detection
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

    def map_and_clean(self, raw_data):
        """Processes multiple tables and normalizes the data."""
        # Handle both single table (list of dicts) and multiple tables (list of list of dicts)
        if raw_data and isinstance(raw_data[0], dict):
            tables = [raw_data]
        else:
            tables = raw_data

        mapped_data = []
        generator = Generator()

        for table in tables:
            if not table:
                continue

            # Identify columns for THIS table
            first_row = table[0]
            standard_cols_mapping = {}
            used_src_cols = set()

            # 1. Explicit mapping
            for target, source in self.mapping.items():
                if source in first_row:
                    standard_cols_mapping[target] = source
                    used_src_cols.add(source)

            # 2. Priority-based heuristic mapping
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Factor', 'ScaleFactor', 'Tag', 'Action']
            for target in detection_order:
                if target in standard_cols_mapping:
                    continue
                patterns = self.COLUMN_MAPPING.get(target, [target.lower()])
                for src_col in first_row.keys():
                    if src_col in used_src_cols:
                        continue
                    src_lower = str(src_col).lower()
                    if any(p in src_lower for p in patterns):
                        standard_cols_mapping[target] = src_col
                        used_src_cols.add(src_col)
                        break

            # Process rows in table
            for row in table:
                new_row = {}
                for target, source in standard_cols_mapping.items():
                    new_row[target] = row.get(source)

                # Include unmapped columns
                for k, v in row.items():
                    if k not in used_src_cols:
                        new_row[k] = v

                # Skip if no Name or Address
                if not new_row.get('Name') and not new_row.get('Address'):
                    continue

                # Normalize Address
                if new_row.get('Address'):
                    addr = str(new_row['Address']).strip()
                    if '_' in addr:
                        parts = addr.split('_')
                        # Support Address_Length and Address_Start_Bit
                        norm_parts = [generator.normalize_address_val(p) for p in parts]
                        new_row['Address'] = '_'.join(norm_parts)
                    else:
                        new_row['Address'] = generator.normalize_address_val(addr)

                # Normalize Type and Action via Generator
                if 'Type' in new_row:
                    new_row['Type'] = generator.normalize_type(new_row['Type'])
                if 'Action' in new_row:
                    new_row['Action'] = generator.normalize_action(new_row['Action'])

                # Handle Factor fractions
                if 'Factor' in new_row:
                    val = str(new_row['Factor'])
                    if '/' in val:
                        try:
                            parts = val.split('/')
                            new_row['Factor'] = str(float(parts[0]) / float(parts[1]))
                        except (ValueError, ZeroDivisionError):
                            pass

                # Defaults
                if not new_row.get('RegisterType'):
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
