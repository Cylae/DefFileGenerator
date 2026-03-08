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
    from defusedxml import ElementTree as ET
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
        wb = openpyxl.load_workbook(filepath, data_only=True)

        sheets_to_process = [sheet_name] if sheet_name else wb.sheetnames
        all_data = []

        for name in sheets_to_process:
            if name not in wb.sheetnames:
                logging.warning(f"Sheet '{name}' not found.")
                continue
            ws = wb[name]
            rows = list(ws.rows)
            if not rows: continue

            headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]
            sheet_data = []
            for row in rows[1:]:
                row_data = {headers[i]: cell.value for i, cell in enumerate(row) if i < len(headers)}
                sheet_data.append(row_data)
            all_data.append(sheet_data)

        return all_data

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        logging.info(f"Extracting from PDF: {filepath}")
        all_tables = []
        with pdfplumber.open(filepath) as pdf:
            target_pages = pdf.pages if pages is None else [pdf.pages[i-1] for i in (pages if isinstance(pages, list) else [pages])]
            for page in target_pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2: continue
                    headers = [str(c).replace('\n', ' ').strip() if c else "" for c in table[0]]
                    table_data = []
                    for row in table[1:]:
                        row_dict = {headers[i]: str(row[i]).replace('\n', ' ').strip() if i < len(row) and row[i] else "" for i in range(len(headers))}
                        table_data.append(row_dict)
                    all_tables.append(table_data)
        return all_tables

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        # Detect encoding
        encoding = 'utf-8-sig'
        try:
            with open(filepath, 'rb') as f:
                raw = f.read(4)
                if raw.startswith(b'\xff\xfe') or raw.startswith(b'\xfe\xff'):
                    encoding = 'utf-16'
        except Exception: pass

        try:
            with open(filepath, 'r', encoding=encoding) as f:
                content = f.read()
                sniffer = csv.Sniffer()
                dialect = sniffer.sniff(content[:2048], delimiters=",;\t")
                f.seek(0)
                reader = csv.DictReader(f, dialect=dialect)
                return [list(reader)]
        except Exception as e:
            logging.error(f"Error reading CSV: {e}")
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
            # Validate with defusedxml before pandas reads it
            ET.fromstring(content)
            df = pd.read_xml(io.BytesIO(content))
            return [df.to_dict(orient='records')]
        except Exception as e:
            logging.error(f"Error reading XML: {e}")
            return []

    def map_and_clean(self, tables):
        if not tables: return []
        # Ensure we're working with a list of tables
        if isinstance(tables, dict) or (isinstance(tables, list) and tables and isinstance(tables[0], dict)):
            tables = [tables]

        final_mapped_data = []
        for table in tables:
            if not table: continue

            first_row = table[0]
            col_mapping = {}
            used_src_cols = set()

            # Explicit mapping
            for target, source in self.mapping.items():
                if source in first_row:
                    col_mapping[target] = source
                    used_src_cols.add(source)

            # Heuristic mapping
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag', 'Factor', 'ScaleFactor']
            for target in detection_order:
                if target in col_mapping: continue
                for src_col in first_row.keys():
                    if src_col in used_src_cols: continue
                    src_lower = str(src_col).lower()
                    if any(p in src_lower for p in self.COLUMN_MAPPING.get(target, [])):
                        col_mapping[target] = src_col
                        used_src_cols.add(src_col)
                        break

            for row in table:
                new_row = {target: row[src] for target, src in col_mapping.items() if src in row}

                # Mandatory Name check
                if not new_row.get('Name'): continue

                # Address cleaning/complex patterns
                addr = str(new_row.get('Address', '')).strip()
                if addr:
                    # Support Address_Length and Address_Start_Bit
                    # Also handles hex conversion via Generator
                    parts = re.split(r'[_/]', addr)
                    norm_parts = [self.generator.normalize_address_val(p) for p in parts]
                    new_row['Address'] = '_'.join(norm_parts)

                # Factor cleaning (fractions)
                factor = str(new_row.get('Factor', '1')).strip()
                if '/' in factor:
                    try:
                        num, den = factor.split('/')
                        new_row['Factor'] = str(float(num) / float(den))
                    except (ValueError, ZeroDivisionError):
                        new_row['Factor'] = '1'

                # Type normalization
                new_row['Type'] = self.normalize_type(new_row.get('Type', 'U16'))

                if 'RegisterType' not in new_row:
                    new_row['RegisterType'] = 'Holding Register'

                final_mapped_data.append(new_row)

        return final_mapped_data

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(description='Extract registers from documentation.')
    parser.add_argument('input_file', help='Source file (PDF, Excel, CSV, XML)')
    parser.add_argument('-o', '--output', help='Output CSV')
    parser.add_argument('--mapping', help='Mapping JSON')
    parser.add_argument('--sheet', help='Excel sheet')
    parser.add_argument('--pages', help='PDF pages')
    args = parser.parse_args()

    mapping = {}
    if args.mapping:
        with open(args.mapping, 'r') as f:
            mapping = json.load(f)

    extractor = Extractor(mapping)
    ext = os.path.splitext(args.input_file)[1].lower()

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xls']:
        data = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = [int(p.strip()) for p in args.pages.split(',')] if args.pages else None
        data = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv':
        data = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml':
        data = extractor.extract_from_xml(args.input_file)
    else:
        logging.error(f"Unsupported extension: {ext}")
        sys.exit(1)

    mapped = extractor.map_and_clean(data)

    output = args.output if args.output else sys.stdout
    fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']

    with (open(output, 'w', newline='', encoding='utf-8') if isinstance(output, str) else output) as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(mapped)

if __name__ == "__main__":
    main()
