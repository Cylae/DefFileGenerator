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
        'Tag': ['tag'],
        'ScaleFactor': ['scalefactor']
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
        sheets = [sheet_name] if sheet_name else wb.sheetnames

        all_tables = []
        for sname in sheets:
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
            if pages is None:
                target_pages = pdf.pages
            else:
                target_pages = []
                if isinstance(pages, int):
                    target_pages = [pdf.pages[pages-1]]
                elif isinstance(pages, list):
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
        for delimiter in [',', ';', '\t']:
            try:
                with open(filepath, 'r', encoding='utf-8-sig') as f:
                    # Check for delimiter by reading first line
                    first_line = f.readline()
                    if delimiter not in first_line:
                        continue
                    f.seek(0)
                    reader = csv.DictReader(f, delimiter=delimiter)
                    rows = list(reader)
                    if rows and len(reader.fieldnames) > 1:
                        # Normalize headers: strip whitespace
                        headers = [h.strip() for h in reader.fieldnames]
                        normalized_rows = []
                        for row in rows:
                            normalized_rows.append({h.strip(): v for h, v in row.items()})
                        return [normalized_rows]
            except Exception:
                continue
        return []

    def extract_from_xml(self, filepath):
        logging.info(f"Extracting from XML: {filepath}")
        if HAS_PANDAS:
            try:
                df = pd.read_xml(filepath)
                return [df.to_dict(orient='records')]
            except Exception as e:
                logging.debug(f"Pandas XML extraction failed: {e}. Trying etree.")

        # Fallback to etree/manual if pandas not present or failed
        import xml.etree.ElementTree as ET
        try:
            tree = ET.parse(filepath)
            root = tree.getroot()
            # Simple heuristic: find all elements that have children (rows)
            all_data = []
            for child in root:
                row_data = {}
                for subchild in child:
                    row_data[subchild.tag] = subchild.text
                if row_data:
                    all_data.append(row_data)
            if all_data:
                return [all_data]
        except Exception as e:
            logging.error(f"XML extraction failed: {e}")
        return []

    def map_and_clean(self, raw_tables):
        """
        Processes a list of tables (each a list of dicts).
        Returns a single list of flattened, cleaned rows.
        """
        if not raw_tables:
            return []

        # Handle case where single table is passed instead of list of tables
        if isinstance(raw_tables, list) and len(raw_tables) > 0 and isinstance(raw_tables[0], dict):
            raw_tables = [raw_tables]

        all_mapped_data = []

        for table in raw_tables:
            if not table:
                continue

            first_row = table[0]
            col_map = {}
            assigned_src_cols = set()

            # Explicitly mapped columns from config
            for target, source in self.mapping.items():
                if source in first_row:
                    col_map[target] = source
                    assigned_src_cols.add(source)

            # Fuzzy match for standard columns
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Scale', 'Action', 'Tag', 'ScaleFactor']
            for target in detection_order:
                if target in col_map:
                    continue
                for col in first_row.keys():
                    if col in assigned_src_cols:
                        continue
                    col_lower = str(col).lower()
                    for pattern in self.COLUMN_MAPPING[target]:
                        if pattern in col_lower:
                            col_map[target] = col
                            assigned_src_cols.add(col)
                            break
                    if target in col_map:
                        break

            if 'Address' not in col_map and 'Name' not in col_map:
                continue

            for row in table:
                addr_raw = row.get(col_map.get('Address')) if 'Address' in col_map else None
                name_raw = row.get(col_map.get('Name')) if 'Name' in col_map else None

                if addr_raw is None and name_raw is None:
                    continue

                # Skip header-like rows
                if str(addr_raw).lower() in self.COLUMN_MAPPING['Address'] or str(name_raw).lower() in self.COLUMN_MAPPING['Name']:
                    continue

                new_row = {}
                # Map identified columns
                for target, source in col_map.items():
                    new_row[target] = row.get(source)

                # Fill in other columns
                for k, v in row.items():
                    if k not in assigned_src_cols and k not in new_row:
                        new_row[k] = v

                # Standardize Address
                if 'Address' in new_row and new_row['Address'] is not None:
                    addr = str(new_row['Address']).strip()
                    if '_' in addr:
                        parts = addr.split('_')
                        norm_parts = [self.generator.normalize_address_val(p) for p in parts]
                        new_row['Address'] = '_'.join(norm_parts)
                    else:
                        new_row['Address'] = self.generator.normalize_address_val(addr)

                # Standardize Type
                if 'Type' in new_row:
                    new_row['Type'] = self.normalize_type(new_row['Type'])

                # Standardize Action
                if 'Action' in new_row:
                    new_row['Action'] = self.normalize_action(new_row['Action'])

                if 'Name' not in new_row or not new_row['Name']:
                    if 'Address' in new_row and new_row['Address']:
                        new_row['Name'] = f"Register {new_row['Address']}"
                    else:
                        continue

                if 'RegisterType' not in new_row:
                    new_row['RegisterType'] = 'Holding Register'

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
    parser.add_argument('--address-offset', type=int, default=0, help='Value to subtract from all addresses.')

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
    if args.address_offset:
        extractor.generator.address_offset = args.address_offset

    ext = os.path.splitext(args.input_file)[1].lower()
    raw_tables = []

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xls']:
        raw_tables = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = None
        if args.pages:
            pages = [int(p.strip()) for p in args.pages.split(',')]
        raw_tables = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv':
        raw_tables = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml':
        raw_tables = extractor.extract_from_xml(args.input_file)
    else:
        logging.error(f"Unsupported file extension: {ext}")
        sys.exit(1)

    if not raw_tables:
        logging.error("No data extracted.")
        sys.exit(1)

    mapped_data = extractor.map_and_clean(raw_tables)

    if not mapped_data:
        logging.error("No data remained after mapping and cleaning.")
        sys.exit(1)

    # Apply final Generator.process_rows to handle Tag generation and final normalization
    # but we want to output a simplified CSV for the 'extract' command,
    # and use 'generate' or 'run' for the full Webdyn file.
    # Actually, main.py 'extract' just calls map_and_clean and writes CSV.
    # map_and_clean already handles some normalization.

    # Write to CSV
    output = args.output if args.output else sys.stdout

    fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']

    if isinstance(output, str):
        f = open(output, 'w', newline='', encoding='utf-8')
    else:
        f = output

    writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore', restval='')
    writer.writeheader()
    writer.writerows(mapped_data)

    if isinstance(output, str):
        f.close()
        logging.info(f"Extracted data saved to {args.output}")

if __name__ == "__main__":
    main()
