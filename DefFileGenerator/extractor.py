#!/usr/bin/env python3
import argparse
import csv
import json
import logging
import os
import re
import sys
import io

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
        'ScaleFactor': ['scalefactor'],
        'Tag': ['tag'],
        'Action': ['action', 'access']
    }

    def __init__(self, mapping=None):
        self.mapping = mapping or {}

    def normalize_action(self, action):
        """Action normalization delegated to Generator."""
        return Generator().normalize_action(action)

    def normalize_type(self, dtype):
        """Type normalization delegated to Generator."""
        return Generator().normalize_type(dtype)

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
        logging.info(f"Extracting from Excel: {filepath}")
        wb = openpyxl.load_workbook(filepath, data_only=True)
        all_tables = []

        target_sheets = [sheet_name] if sheet_name else wb.sheetnames
        for sname in target_sheets:
            if sname not in wb.sheetnames:
                continue
            ws = wb[sname]
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
            target_pages = pdf.pages
            if pages:
                target_pages = [pdf.pages[i-1] for i in pages if 0 < i <= len(pdf.pages)]

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
        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                content = f.read(4096)
                f.seek(0)
                # Manual delimiter detection
                delims = [',', ';', '\t']
                best_delim = ','
                max_cols = 0
                for d in delims:
                    cols = content.split('\n')[0].count(d)
                    if cols > max_cols:
                        max_cols = cols
                        best_delim = d

                reader = csv.DictReader(f, delimiter=best_delim)
                table_data = list(reader)
                if table_data:
                    all_tables.append(table_data)
        except Exception as e:
            logging.error(f"Error loading CSV file: {e}")
        return all_tables

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
        """Processes multiple tables and normalizes register data."""
        if not tables:
            return []
        # Support single table for backward compatibility
        if isinstance(tables, list) and tables and isinstance(tables[0], dict):
            tables = [tables]

        final_data = []
        generator = Generator()

        for table in tables:
            if not table:
                continue
            first_row = table[0]
            col_map = {}
            used_src_cols = set()

            # Priority 1: Explicit mapping
            for target, source in self.mapping.items():
                if source in first_row:
                    col_map[target] = source
                    used_src_cols.add(source)

            # Priority 2: Heuristic mapping with mandated priority
            prio_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag', 'Factor', 'ScaleFactor']
            for target in prio_order:
                if target in col_map:
                    continue
                patterns = self.COLUMN_MAPPING.get(target, [target.lower()])
                for src_col in first_row.keys():
                    if src_col in used_src_cols:
                        continue
                    src_lower = str(src_col).lower()
                    if any(p in src_lower for p in patterns):
                        col_map[target] = src_col
                        used_src_cols.add(src_col)
                        break

            for row in table:
                new_row = {}
                for target, source in col_map.items():
                    new_row[target] = row.get(source)

                # Mandated: support fraction parsing for Factor
                factor_val = str(new_row.get('Factor', '1.0'))
                if '/' in factor_val:
                    try:
                        p1, p2 = factor_val.split('/')
                        new_row['Factor'] = str(float(p1) / float(p2))
                    except (ValueError, ZeroDivisionError):
                        new_row['Factor'] = '1.0'

                # Mandated: support complex address patterns
                addr = str(new_row.get('Address', '')).strip()
                if not addr:
                    continue
                # Normalize via Generator
                if '_' in addr:
                    parts = addr.split('_')
                    norm_parts = [generator.normalize_address_val(p) for p in parts]
                    new_row['Address'] = '_'.join(norm_parts)
                else:
                    new_row['Address'] = generator.normalize_address_val(addr)

                if not new_row.get('Name'):
                    new_row['Name'] = f"Reg {new_row['Address']}"

                # Data type and Action normalization is delegated to Generator.process_rows,
                # but we perform it here as well to ensure the simplified CSV is clean.
                if 'Type' in new_row:
                    new_row['Type'] = self.normalize_type(new_row['Type'])
                if 'Action' in new_row:
                    new_row['Action'] = self.normalize_action(new_row['Action'])

                final_data.append(new_row)
        return final_data

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(description='Extract registers from documentation.')
    parser.add_argument('input_file', help='PDF, Excel, CSV, or XML file.')
    parser.add_argument('-o', '--output', help='Output CSV file.')
    parser.add_argument('--mapping', help='Mapping JSON file.')
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

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xls']:
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
        logging.error("No tables extracted.")
        sys.exit(1)

    mapped_data = extractor.map_and_clean(tables)
    if not mapped_data:
        logging.error("No data after mapping.")
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
        logging.info(f"Saved to {args.output}")

if __name__ == "__main__":
    main()
