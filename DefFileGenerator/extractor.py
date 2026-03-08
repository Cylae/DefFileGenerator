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

try:
    import defusedxml
    from defusedxml.ElementTree import parse as parse_xml
    HAS_DEFUSEDXML = True
except ImportError:
    HAS_DEFUSEDXML = False

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

    def normalize_type(self, t):
        return self.generator.normalize_type(t)

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
        logging.info(f"Extracting from Excel: {filepath}")
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            if sheet_name:
                if sheet_name not in wb.sheetnames:
                    logging.error(f"Sheet '{sheet_name}' not found in {filepath}")
                    return []
                sheets = [wb[sheet_name]]
            else:
                sheets = wb.worksheets

            all_data = []
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
                    all_data.append(data)
            return all_data
        except Exception as e:
            logging.error(f"Error loading Excel file: {e}")
            return []

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        logging.info(f"Extracting from PDF: {filepath}")
        all_data = []
        try:
            with pdfplumber.open(filepath) as pdf:
                if pages is None:
                    target_pages = pdf.pages
                else:
                    target_pages = [pdf.pages[i-1] for i in pages if 0 < i <= len(pdf.pages)]

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
                            all_data.append(data)
            return all_data
        except Exception as e:
            logging.error(f"Error loading PDF file: {e}")
            return []

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        try:
            # Detect encoding and delimiter
            with open(filepath, 'rb') as f:
                raw_content = f.read(4096)
                encoding = 'utf-16' if raw_content.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'

            with open(filepath, 'r', encoding=encoding) as f:
                sample = f.read(4096)
                f.seek(0)
                delimiter = ','
                for d in [',', ';', '\t']:
                    if d in sample:
                        delimiter = d
                        break
                reader = csv.DictReader(f, delimiter=delimiter)
                data = list(reader)
                return [data] if data else []
        except Exception as e:
            logging.error(f"Error loading CSV file: {e}")
            return []

    def extract_from_xml(self, filepath):
        if not HAS_DEFUSEDXML:
            logging.error("defusedxml is required for secure XML extraction.")
            return []
        if not HAS_PANDAS:
            logging.error("pandas is required for XML extraction.")
            return []
        logging.info(f"Extracting from XML: {filepath}")
        try:
            with open(filepath, 'rb') as f:
                content = f.read()
            # Validate with defusedxml before parsing with pandas
            defusedxml.ElementTree.fromstring(content)
            df = pd.read_xml(io.BytesIO(content))
            return [df.to_dict(orient='records')]
        except Exception as e:
            logging.error(f"Error loading XML file: {e}")
            return []

    def map_and_clean(self, raw_tables):
        """Processes one or more tables of raw data."""
        if not raw_tables:
            return []

        # Support both a single table or a list of tables
        if isinstance(raw_tables, list) and raw_tables and not isinstance(raw_tables[0], dict):
            tables = raw_tables
        else:
            tables = [raw_tables]

        all_mapped_data = []

        for table in tables:
            if not table:
                continue

            first_row = table[0]
            col_map = {}
            used_src_cols = set()

            # 1. Explicit mapping from config
            for target, source in self.mapping.items():
                if source in first_row:
                    col_map[target] = source
                    used_src_cols.add(source)

            # 2. Heuristic mapping
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag', 'Factor', 'ScaleFactor']
            for target in detection_order:
                if target in col_map:
                    continue
                patterns = self.COLUMN_MAPPING.get(target, [])
                for col in first_row.keys():
                    if col in used_src_cols:
                        continue
                    col_lower = str(col).lower()
                    if any(p in col_lower for p in patterns):
                        col_map[target] = col
                        used_src_cols.add(col)
                        break

            if 'Address' not in col_map and 'Name' not in col_map:
                continue

            for row in table:
                new_row = {}
                for target, source in col_map.items():
                    new_row[target] = row.get(source)

                # Name is mandatory
                if not new_row.get('Name'):
                    continue

                # Normalize Address
                addr = str(new_row.get('Address', '')).strip()
                if addr:
                    if '_' in addr:
                        parts = addr.split('_')
                        new_row['Address'] = '_'.join([self.generator.normalize_address_val(p) for p in parts])
                    else:
                        new_row['Address'] = self.generator.normalize_address_val(addr)

                # Normalize Type
                new_row['Type'] = self.normalize_type(new_row.get('Type'))

                # Normalize Factor (fractions)
                factor = str(new_row.get('Factor', '1'))
                if '/' in factor:
                    try:
                        p = factor.split('/')
                        new_row['Factor'] = str(float(p[0]) / float(p[1]))
                    except:
                        new_row['Factor'] = '1'

                if 'RegisterType' not in new_row:
                    new_row['RegisterType'] = 'Holding Register'

                all_mapped_data.append(new_row)

        return all_mapped_data

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

    parser = argparse.ArgumentParser(description='Extract register information from manufacturer documentation.')
    parser.add_argument('input_file', help='Source file (PDF, Excel, CSV, XML).')
    parser.add_argument('-o', '--output', help='Output CSV.')
    parser.add_argument('--mapping', help='JSON file containing column mapping.')
    parser.add_argument('--sheet', help='Excel sheet name.')
    parser.add_argument('--pages', help='PDF pages.')

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
    raw_tables = []

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw_tables = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = [int(p.strip()) for p in args.pages.split(',')] if args.pages else None
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
