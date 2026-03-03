#!/usr/bin/env python3
import argparse
import csv
import json
import logging
import os
import re
import sys
import math

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

                    headers = [str(c).replace('\n', ' ').strip() if c else f"Col{i}" for i, c in enumerate(table[0])]
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
        all_tables = []
        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                # Detect delimiter
                sample = f.read(4096)
                f.seek(0)
                delimiter = ','
                for d in [',', ';', '\t']:
                    if d in sample:
                        delimiter = d
                        break

                reader = csv.DictReader(f, delimiter=delimiter)
                data = list(reader)
                if data:
                    all_tables.append(data)
        except Exception as e:
            logging.error(f"Error extracting from CSV: {e}")
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
            logging.debug(f"Pandas read_xml failed, trying etree: {e}")
            try:
                import xml.etree.ElementTree as ET
                tree = ET.parse(filepath)
                root = tree.getroot()
                data = []
                for child in root:
                    row_data = {}
                    for subchild in child:
                        row_data[subchild.tag] = subchild.text
                    data.append(row_data)
                return [data] if data else []
            except Exception as e2:
                logging.error(f"Error extracting from XML: {e2}")
                return []

    def map_and_clean(self, tables):
        """Processes extracted tables into a list of mapped registers."""
        if not tables:
            return []

        # Handle both single table or list of tables for backward compatibility
        if isinstance(tables, list) and tables and isinstance(tables[0], dict):
            tables = [tables]

        all_mapped_data = []
        generator = Generator()

        for raw_data in tables:
            if not raw_data:
                continue

            first_row = raw_data[0]
            src_cols = list(first_row.keys())

            # Identify columns
            col_map = {}
            used_src_cols = set()

            # Explicitly mapped columns from config
            for target, source in self.mapping.items():
                if source in src_cols:
                    col_map[target] = source
                    used_src_cols.add(source)

            # Heuristic match
            # Priority order for detection
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Factor', 'ScaleFactor', 'Tag', 'Action']
            for target in detection_order:
                if target in col_map:
                    continue
                for col in src_cols:
                    if col in used_src_cols:
                        continue
                    col_lower = str(col).lower()
                    for pattern in self.COLUMN_MAPPING.get(target, []):
                        if pattern in col_lower:
                            col_map[target] = col
                            used_src_cols.add(col)
                            break
                    if target in col_map:
                        break

            if 'Address' not in col_map and 'Name' not in col_map:
                continue

            for row in raw_data:
                new_row = {}
                for target, source in col_map.items():
                    val = row.get(source)
                    if val is not None:
                        new_row[target] = str(val).strip()

                # Add extra columns that were not mapped
                for k, v in row.items():
                    if k not in used_src_cols:
                        new_row[k] = v

                # Basic validation: must have Name or Address
                if not new_row.get('Name') and not new_row.get('Address'):
                    continue

                # Normalize Address (initial cleaning, Generator will do the rest)
                if 'Address' in new_row and new_row['Address']:
                    addr = str(new_row['Address']).strip()
                    # Handle "40,001"
                    if ',' in addr and '.' not in addr:
                        addr = addr.replace(',', '')
                    new_row['Address'] = addr

                # Default RegisterType
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
    parser.add_argument('--pages', help='PDF pages to extract from.')

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
