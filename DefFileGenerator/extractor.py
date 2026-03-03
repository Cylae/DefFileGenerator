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
    # Column naming heuristics
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

    def normalize_type(self, t):
        # Delegate to Generator
        return Generator().normalize_type(t)

    def normalize_action(self, action):
        # Delegate to Generator
        return Generator().normalize_action(action)

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
        logging.info(f"Extracting from Excel: {filepath}")
        wb = openpyxl.load_workbook(filepath, data_only=True)

        all_tables = []
        sheets_to_process = [sheet_name] if sheet_name else wb.sheetnames

        for name in sheets_to_process:
            if name not in wb.sheetnames:
                logging.error(f"Sheet '{name}' not found in {filepath}")
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

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        logging.info(f"Extracting from PDF: {filepath}")
        all_tables_data = []
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
                    headers = [str(c).replace('\n', ' ').strip() if c else f"Col{i}" for i, c in enumerate(table[0])]

                    table_data = []
                    for row in table[1:]:
                        row_data = {}
                        for i, cell in enumerate(row):
                            if i < len(headers):
                                val = str(cell).replace('\n', ' ').strip() if cell else ""
                                row_data[headers[i]] = val
                        table_data.append(row_data)
                    all_tables_data.append(table_data)
        return all_tables_data

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        # Detect delimiter
        delimiter = ','
        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                sample = f.read(2048)
                if '\t' in sample: delimiter = '\t'
                elif ';' in sample: delimiter = ';'
                elif ',' in sample: delimiter = ','
        except Exception:
            pass

        data = []
        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f, delimiter=delimiter)
                for row in reader:
                    data.append({k.strip() if k else "": v for k, v in row.items()})
        except Exception as e:
            logging.error(f"Error reading CSV: {e}")
        return [data] if data else []

    def extract_from_xml(self, filepath):
        logging.info(f"Extracting from XML: {filepath}")
        try:
            import pandas as pd
            try:
                df = pd.read_xml(filepath)
                return [df.to_dict(orient='records')]
            except Exception as e:
                logging.debug(f"read_xml failed: {e}, trying etree parser")
                df = pd.read_xml(filepath, parser='etree')
                return [df.to_dict(orient='records')]
        except ImportError:
            logging.error("pandas and lxml/etree required for XML extraction.")
            return []
        except Exception as e:
            logging.error(f"Error reading XML: {e}")
            return []

    def map_and_clean(self, raw_tables):
        """
        Maps source columns to standard names and cleans the data.
        Accepts a list of tables (list of lists of dicts) or a single table.
        """
        if not raw_tables:
            return []

        # Wrap in a list if a single table was passed for backward compatibility
        if raw_tables and isinstance(raw_tables, list) and len(raw_tables) > 0 and isinstance(raw_tables[0], dict):
            raw_tables = [raw_tables]

        all_mapped_data = []
        generator = Generator()

        for table in raw_tables:
            if not table:
                continue

            # Identify standard columns for this table
            first_row = table[0]
            table_col_mapping = {}
            used_src_cols = set()

            # 1. Explicitly mapped columns from config
            for target, source in self.mapping.items():
                if source in first_row:
                    table_col_mapping[target] = source
                    used_src_cols.add(source)

            # 2. Priority-based heuristic matching for standard columns
            priority_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag', 'Factor', 'ScaleFactor']
            for target in priority_order:
                if target in table_col_mapping:
                    continue

                for src_col in first_row.keys():
                    if src_col in used_src_cols:
                        continue

                    src_col_lower = str(src_col).lower()
                    for pattern in self.COLUMN_MAPPING.get(target, []):
                        if pattern in src_col_lower:
                            table_col_mapping[target] = src_col
                            used_src_cols.add(src_col)
                            break
                    if target in table_col_mapping:
                        break

            if 'Address' not in table_col_mapping and 'Name' not in table_col_mapping:
                 continue

            for row in table:
                new_row = {}
                for target, source in table_col_mapping.items():
                    val = row.get(source)
                    if val is not None:
                        new_row[target] = val

                # Clean values
                if 'Name' not in new_row or not str(new_row.get('Name', '')).strip():
                    if 'Address' in new_row:
                        new_row['Name'] = f"Register {new_row['Address']}"
                    else:
                        continue

                # Normalize Address
                if 'Address' in new_row:
                    addr = str(new_row['Address']).strip()
                    if '_' in addr:
                        parts = addr.split('_')
                        norm_parts = [generator.normalize_address_val(p) for p in parts]
                        new_row['Address'] = '_'.join(norm_parts)
                    else:
                        new_row['Address'] = generator.normalize_address_val(addr)

                # Normalize other fields via Generator where applicable
                # Note: We leave heavy type normalization for the Generator.process_rows
                # but we can do a light pass here for consistency in the intermediate CSV.
                if 'Type' in new_row:
                     new_row['Type'] = self.normalize_type(new_row['Type'])
                if 'Action' in new_row:
                     new_row['Action'] = self.normalize_action(new_row['Action'])

                if 'RegisterType' not in new_row:
                    new_row['RegisterType'] = 'Holding Register'

                # Add factor normalization (e.g. 1/10 -> 0.1)
                if 'Factor' in new_row:
                    factor = str(new_row['Factor'])
                    if '/' in factor:
                        try:
                            parts = factor.split('/')
                            new_row['Factor'] = str(float(parts[0]) / float(parts[1]))
                        except (ValueError, ZeroDivisionError):
                            pass

                all_mapped_data.append(new_row)

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
