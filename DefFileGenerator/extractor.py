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
        'Action': ['action', 'access'],
        'Tag': ['tag'],
        'ScaleFactor': ['scalefactor']
    }

    def __init__(self, mapping=None):
        self.mapping = mapping or {}

    def extract_from_csv(self, filepath):
        """Extracts data from CSV with manual delimiter detection."""
        logging.info(f"Extracting from CSV: {filepath}")
        data = []
        delimiters = [',', ';', '\t']

        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                sample = f.read(1024)
                f.seek(0)

                detected_sep = ','
                for sep in delimiters:
                    if sep in sample:
                        detected_sep = sep
                        break

                reader = csv.DictReader(f, delimiter=detected_sep)
                data = [list(reader)]
        except Exception as e:
            logging.error(f"Error extracting from CSV: {e}")

        return data

    def extract_from_excel(self, filepath, sheet_name=None):
        """Extracts data from Excel, returning a list of tables."""
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []

        logging.info(f"Extracting from Excel: {filepath}")
        wb = openpyxl.load_workbook(filepath, data_only=True)
        all_tables = []

        sheets = [sheet_name] if sheet_name else wb.sheetnames
        for name in sheets:
            if name not in wb.sheetnames:
                continue
            ws = wb[name]
            rows = list(ws.rows)
            if not rows:
                continue

            headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]
            table_data = []
            for row in rows[1:]:
                row_dict = {}
                for i, cell in enumerate(row):
                    if i < len(headers):
                        row_dict[headers[i]] = cell.value
                table_data.append(row_dict)
            all_tables.append(table_data)

        return all_tables

    def extract_from_pdf(self, filepath, pages=None):
        """Extracts data from PDF tables, returning a list of tables."""
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []

        logging.info(f"Extracting from PDF: {filepath}")
        all_tables = []
        with pdfplumber.open(filepath) as pdf:
            target_pages = pdf.pages
            if pages:
                if isinstance(pages, int):
                    target_pages = [pdf.pages[pages-1]]
                elif isinstance(pages, list):
                    target_pages = [pdf.pages[i-1] for i in pages if 0 < i <= len(pdf.pages)]

            for page in target_pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:
                        continue

                    headers = [str(c).replace('\n', ' ').strip() if c else "" for c in table[0]]
                    table_data = []
                    for row in table[1:]:
                        row_dict = {}
                        for i, cell in enumerate(row):
                            if i < len(headers):
                                row_dict[headers[i]] = str(cell).replace('\n', ' ').strip() if cell else ""
                        table_data.append(row_dict)
                    all_tables.append(table_data)

        return all_tables

    def extract_from_xml(self, filepath):
        """Extracts data from XML using pandas or etree."""
        if not HAS_PANDAS:
            logging.error("pandas is required for XML extraction.")
            return []

        logging.info(f"Extracting from XML: {filepath}")
        try:
            df = pd.read_xml(filepath)
            return [df.to_dict(orient='records')]
        except Exception as e:
            logging.debug(f"Pandas read_xml failed, trying etree parser: {e}")
            try:
                df = pd.read_xml(filepath, parser='etree')
                return [df.to_dict(orient='records')]
            except Exception as e2:
                logging.error(f"Error extracting from XML: {e2}")
                return []

    def map_and_clean(self, raw_data):
        """Maps source columns to standard columns and cleans data."""
        if not raw_data:
            return []

        # Handle single table or list of tables
        tables = raw_data if isinstance(raw_data[0], list) else [raw_data]
        mapped_data = []
        generator = Generator()

        for table in tables:
            if not table:
                continue

            first_row = table[0]
            src_cols = list(first_row.keys())
            col_map = {}
            used_src_cols = set()

            # 1. Explicit mapping
            for target, source in self.mapping.items():
                if source in src_cols:
                    col_map[target] = source
                    used_src_cols.add(source)

            # 2. Priority-based heuristic mapping
            priority_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag', 'Factor', 'ScaleFactor']
            for target in priority_order:
                if target in col_map:
                    continue

                patterns = self.COLUMN_MAPPING.get(target, [target.lower()])
                for src in src_cols:
                    if src in used_src_cols:
                        continue

                    src_lower = str(src).lower()
                    if any(p in src_lower for p in patterns):
                        col_map[target] = src
                        used_src_cols.add(src)
                        break

            for row in table:
                new_row = {}
                for target, source in col_map.items():
                    val = row.get(source)
                    new_row[target] = val

                if not new_row.get('Name') and not new_row.get('Address'):
                    continue

                # Support Address_Length and Address_Start_Bit patterns
                addr = str(new_row.get('Address', '')).strip()
                if addr:
                    # Basic cleaning for complex patterns
                    addr = re.sub(r'\s+', ' ', addr)
                    new_row['Address'] = addr

                # Fraction parsing for Factor
                factor = new_row.get('Factor')
                if factor and isinstance(factor, str) and '/' in factor:
                    try:
                        p1, p2 = factor.split('/')
                        new_row['Factor'] = str(float(p1) / float(p2))
                    except (ValueError, ZeroDivisionError):
                        pass

                # Default RegisterType
                if not new_row.get('RegisterType'):
                    new_row['RegisterType'] = 'Holding Register'

                mapped_data.append(new_row)

        return mapped_data

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(description='Extract register information from documentation.')
    parser.add_argument('input_file', help='Source file (PDF, Excel, CSV, XML)')
    parser.add_argument('-o', '--output', help='Output CSV file')
    parser.add_argument('--mapping', help='JSON mapping file')
    parser.add_argument('--sheet', help='Excel sheet name')
    parser.add_argument('--pages', help='PDF pages')

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

    if ext == '.csv':
        raw_data = extractor.extract_from_csv(args.input_file)
    elif ext in ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xls']:
        raw_data = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = [int(p.strip()) for p in args.pages.split(',')] if args.pages else None
        raw_data = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.xml':
        raw_data = extractor.extract_from_xml(args.input_file)
    else:
        logging.error(f"Unsupported extension: {ext}")
        sys.exit(1)

    if not raw_data:
        logging.error("No data extracted.")
        sys.exit(1)

    mapped_data = extractor.map_and_clean(raw_data)

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
