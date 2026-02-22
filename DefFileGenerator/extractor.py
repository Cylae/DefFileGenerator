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

    TYPE_MAPPING = {
        'uint16': 'U16',
        'int16': 'I16',
        'uint32': 'U32',
        'int32': 'I32',
        'uint64': 'U64',
        'int64': 'I64',
        'float': 'F32',
        'f32': 'F32',
        'float32': 'F32',
        'double': 'F64',
        'f64': 'F64',
        'float64': 'F64',
        'string': 'STRING',
        'bits': 'BITS'
    }

    def __init__(self, mapping=None):
        self.mapping = mapping or {}

    def normalize_type(self, t):
        if not t:
            return 'U16'
        t_str = str(t).lower().strip()
        # Remove common extra words and spaces
        t_str = t_str.replace('unsigned ', 'u').replace('signed ', 'i').replace(' ', '')

        # Check in mapping
        for key, val in self.TYPE_MAPPING.items():
            if key in t_str:
                return val

        # Check for patterns like Uint16, Int32, uint16, int32
        match = re.match(r'^(u|i|uint|int)(\d+)$', t_str)
        if match:
            raw_prefix = match.group(1).lower()
            prefix = 'U' if raw_prefix.startswith('u') else 'I'
            bits = match.group(2)
            return f"{prefix}{bits}"

        # Regex based cleaning
        t_clean = re.sub(r'[^a-z0-9_]+', '', t_str)
        return t_clean.upper() if t_clean else 'U16'

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
        logging.info(f"Extracting from Excel: {filepath}")
        wb = openpyxl.load_workbook(filepath, data_only=True)

        tables = []
        sheets = [sheet_name] if sheet_name else wb.sheetnames

        for name in sheets:
            if name not in wb.sheetnames:
                logging.warning(f"Sheet '{name}' not found.")
                continue
            ws = wb[name]
            rows = list(ws.rows)
            if len(rows) < 2:
                continue

            headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]
            sheet_data = []
            for row in rows[1:]:
                row_data = {}
                for i, cell in enumerate(row):
                    if i < len(headers):
                        row_data[headers[i]] = cell.value
                sheet_data.append(row_data)
            tables.append(sheet_data)
        return tables

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        logging.info(f"Extracting from PDF: {filepath}")
        tables = []
        with pdfplumber.open(filepath) as pdf:
            if pages is None:
                target_pages = pdf.pages
            else:
                target_pages = []
                for p_num in pages:
                    if 1 <= p_num <= len(pdf.pages):
                        target_pages.append(pdf.pages[p_num-1])

            for page in target_pages:
                extracted_tables = page.extract_tables()
                for table in extracted_tables:
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
                    tables.append(table_data)
        return tables

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        tables = []
        # Detect delimiter
        delimiter = ','
        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                content = f.read(2048)
                if ';' in content: delimiter = ';'
                elif '\t' in content: delimiter = '\t'
                f.seek(0)
                reader = csv.DictReader(f, delimiter=delimiter)
                data = list(reader)
                if data:
                    tables.append(data)
        except Exception as e:
            logging.error(f"Error reading CSV: {e}")
        return tables

    def extract_from_xml(self, filepath):
        logging.info(f"Extracting from XML: {filepath}")
        try:
            import pandas as pd
            df = pd.read_xml(filepath)
            return [df.to_dict(orient='records')]
        except ImportError:
            logging.error("pandas and lxml are required for XML extraction.")
        except Exception as e:
            logging.error(f"Error reading XML: {e}")
        return []

    def map_and_clean(self, tables):
        """
        Maps raw extracted data to standard definition columns.
        Supports both a single table (list of dicts) or multiple tables (list of lists of dicts).
        """
        if not tables:
            return []

        # Wrap single table in a list for uniform processing
        if tables and isinstance(tables[0], dict):
            tables = [tables]

        all_mapped_data = []
        generator = Generator() # Used for normalization

        for table in tables:
            if not table:
                continue

            # Identify standard columns for THIS table
            first_row = table[0]
            standard_cols_mapping = {}
            used_src_cols = set()

            # Explicitly mapped columns from config
            for target, source in self.mapping.items():
                if source in first_row:
                    standard_cols_mapping[target] = source
                    used_src_cols.add(source)

            # Priority-based heuristic mapping
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Tag', 'Factor', 'ScaleFactor', 'Action']
            for target in detection_order:
                if target in standard_cols_mapping:
                    continue

                # Try to find a match in the table headers
                for src_col in first_row.keys():
                    if src_col in used_src_cols:
                        continue

                    src_col_lower = str(src_col).lower()
                    patterns = self.COLUMN_MAPPING.get(target, [])
                    if any(p in src_col_lower for p in patterns):
                        standard_cols_mapping[target] = src_col
                        used_src_cols.add(src_col)
                        break

            for row in table:
                new_row = {}
                for target, source in standard_cols_mapping.items():
                    new_row[target] = row.get(source)

                # Fill in unmapped columns
                for k, v in row.items():
                    if k not in used_src_cols and k not in new_row:
                        new_row[k] = v

                # Clean Address: use Generator's normalization
                if 'Address' in new_row and new_row['Address'] is not None:
                    addr_raw = str(new_row['Address']).strip()
                    parts = addr_raw.split('_')
                    norm_parts = [generator.normalize_address_val(p) for p in parts]
                    new_row['Address'] = '_'.join(norm_parts)

                # Clean Type
                if 'Type' in new_row:
                    new_row['Type'] = self.normalize_type(new_row['Type'])

                # Skip rows without Name or Address after cleaning
                if not new_row.get('Name') and not new_row.get('Address'):
                    continue

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
    raw_tables = []

    if ext in ['.xlsx', '.xls', '.xlsm', '.xltx', '.xltm']:
        raw_tables = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = None
        if args.pages:
            pages = [int(p.strip()) for p in args.pages.split(',')]
        raw_tables = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv':
        raw_tables = extractor.extract_from_csv(args.input_file)
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
