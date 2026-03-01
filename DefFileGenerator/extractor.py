#!/usr/bin/env python3
import argparse
import csv
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
        'ScaleFactor': ['scalefactor'],
        'Tag': ['tag']
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
        all_tables = []
        with pdfplumber.open(filepath) as pdf:
            if pages is None:
                target_pages = pdf.pages
            else:
                target_pages = []
                for p in (pages if isinstance(pages, list) else [pages]):
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
        # Manual delimiter check
        delimiters = [',', ';', '\t']
        best_delimiter = ','
        max_cols = 0

        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                sample = f.read(2048)
                for d in delimiters:
                    cols = sample.split('\n')[0].count(d)
                    if cols > max_cols:
                        max_cols = cols
                        best_delimiter = d

            with open(filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f, delimiter=best_delimiter)
                data = list(reader)
                return [data] if data else []
        except Exception as e:
            logging.error(f"Error extracting from CSV: {e}")
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
        """Processes a list of tables and returns a flattened list of cleaned rows."""
        # Handle single table input for backward compatibility
        if tables and isinstance(tables[0], dict):
            tables = [tables]

        flattened_data = []

        for table in tables:
            if not table:
                continue

            first_row = table[0]
            col_map = {}
            used_src_cols = set()

            # 1. Explicit mapping
            for target, source in self.mapping.items():
                if source in first_row:
                    col_map[target] = source
                    used_src_cols.add(source)

            # 2. Heuristic mapping
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Factor', 'Action', 'ScaleFactor', 'Tag']
            for target in detection_order:
                if target in col_map:
                    continue
                for src_col in first_row.keys():
                    if src_col in used_src_cols:
                        continue
                    src_col_lower = str(src_col).lower()
                    for pattern in self.COLUMN_MAPPING.get(target, []):
                        if pattern in src_col_lower:
                            col_map[target] = src_col
                            used_src_cols.add(src_col)
                            break
                    if target in col_map:
                        break

            if 'Address' not in col_map and 'Name' not in col_map:
                continue

            for row in table:
                new_row = {}
                for target, source in col_map.items():
                    val = row.get(source)
                    if val is not None:
                        new_row[target] = val

                # Skip empty or header-like rows
                name_val = str(new_row.get('Name', '')).lower()
                addr_val = str(new_row.get('Address', '')).lower()
                if not name_val and not addr_val:
                    continue
                if name_val in self.COLUMN_MAPPING['Name'] or addr_val in self.COLUMN_MAPPING['Address']:
                    continue

                # Normalization delegates to Generator
                if 'Type' in new_row:
                    new_row['Type'] = self.generator.normalize_type(new_row['Type'])
                else:
                    new_row['Type'] = 'U16'

                if 'Address' in new_row:
                    addr = str(new_row['Address']).strip()
                    if '_' in addr:
                        parts = addr.split('_')
                        new_row['Address'] = '_'.join([self.generator.normalize_address_val(p) for p in parts])
                    else:
                        new_row['Address'] = self.generator.normalize_address_val(addr)

                if 'Action' in new_row:
                    new_row['Action'] = self.generator.normalize_action(new_row['Action'])

                if 'Factor' in new_row:
                    # Clean scale (sometimes it's "1/10" or "0.1")
                    scale = str(new_row['Factor'])
                    if '/' in scale:
                        try:
                            parts = scale.split('/')
                            new_row['Factor'] = str(float(parts[0]) / float(parts[1]))
                        except (ValueError, ZeroDivisionError):
                            new_row['Factor'] = '1'

                if 'Name' not in new_row or not str(new_row['Name']).strip():
                    if 'Address' in new_row:
                        new_row['Name'] = f"Register {new_row['Address']}"
                    else:
                        continue

                flattened_data.append(new_row)

        return flattened_data

def main():
    import json
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

    parser = argparse.ArgumentParser(description='Extract register information.')
    parser.add_argument('input_file', help='Path to source file.')
    parser.add_argument('-o', '--output', help='Path to output CSV.')
    parser.add_argument('--mapping', help='JSON mapping file.')
    parser.add_argument('--sheet', help='Excel sheet name.')
    parser.add_argument('--pages', help='PDF pages (comma separated).')

    args = parser.parse_args()

    mapping = {}
    if args.mapping:
        try:
            with open(args.mapping, 'r') as f:
                mapping = json.load(f)
        except Exception as e:
            logging.error(f"Error loading mapping: {e}")
            sys.exit(1)

    extractor = Extractor(mapping)
    ext = os.path.splitext(args.input_file)[1].lower()
    tables = []

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        tables = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = [int(p.strip()) for p in args.pages.split(',')] if args.pages else None
        tables = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv':
        tables = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml':
        tables = extractor.extract_from_xml(args.input_file)
    else:
        logging.error(f"Unsupported extension: {ext}")
        sys.exit(1)

    if not tables:
        logging.error("No data extracted.")
        sys.exit(1)

    mapped_data = extractor.map_and_clean(tables)
    if not mapped_data:
        logging.error("No data remained after mapping.")
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
        logging.info(f"Data saved to {args.output}")

if __name__ == "__main__":
    main()
