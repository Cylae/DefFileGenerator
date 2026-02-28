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
from DefFileGenerator.security_utils import ensure_safe_path

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
        # Default mapping for data types
        self.type_mapping = {
            'uint16': 'U16',
            'int16': 'I16',
            'uint32': 'U32',
            'int32': 'I32',
            'float32': 'F32',
            'float': 'F32',
            'u16': 'U16',
            'i16': 'I16',
            'u32': 'U32',
            'i32': 'I32',
            'f32': 'F32',
            'string': 'STRING',
            'bits': 'BITS'
        }

    def normalize_type(self, t):
        if not t:
            return 'U16'
        t_str = str(t).lower().strip()
        # Remove common extra words and spaces
        t_str = t_str.replace('unsigned ', 'u').replace('signed ', 'i').replace(' ', '')

        if t_str in self.type_mapping:
            return self.type_mapping[t_str]

        # Check for patterns like Uint16, Int32, uint16, int32
        match = re.match(r'^(u|i|uint|int)(\d+)$', t_str)
        if match:
            raw_prefix = match.group(1).lower()
            prefix = 'U' if raw_prefix.startswith('u') else 'I'
            bits = match.group(2)
            return f"{prefix}{bits}"

        return t.upper()

    def normalize_action(self, action):
        if action is None or action == '':
            return '1'
        a = str(action).upper().strip()
        if a in ['R', 'READ']:
            return '4'
        if a in ['RW', 'W', 'WRITE']:
            return '1'
        return a

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []

        safe_path = ensure_safe_path(filepath)
        logging.info(f"Extracting from Excel: {safe_path}")

        wb = openpyxl.load_workbook(safe_path, data_only=True)
        sheets = [sheet_name] if sheet_name else wb.sheetnames

        all_tables = []
        for s_name in sheets:
            if s_name not in wb.sheetnames:
                logging.warning(f"Sheet '{s_name}' not found in {filepath}")
                continue

            ws = wb[s_name]
            rows = list(ws.rows)
            if not rows:
                continue

            headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]
            data = []
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

        safe_path = ensure_safe_path(filepath)
        logging.info(f"Extracting from PDF: {safe_path}")

        all_tables = []
        with pdfplumber.open(safe_path) as pdf:
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
        safe_path = ensure_safe_path(filepath)
        logging.info(f"Extracting from CSV: {safe_path}")

        # Detect delimiter
        delimiter = ','
        try:
            with open(safe_path, 'r', encoding='utf-8-sig') as f:
                content = f.read(2048)
                if ';' in content and content.count(';') > content.count(','):
                    delimiter = ';'
                elif '\t' in content:
                    delimiter = '\t'
        except Exception:
            pass

        data = []
        try:
            with open(safe_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f, delimiter=delimiter)
                for row in reader:
                    data.append({k.strip(): v for k, v in row.items() if k})
        except Exception as e:
            logging.error(f"Error loading CSV file: {e}")
            return []

        return [data] if data else []

    def extract_from_xml(self, filepath):
        if not HAS_PANDAS:
            logging.error("pandas is required for XML extraction.")
            return []

        safe_path = ensure_safe_path(filepath)
        logging.info(f"Extracting from XML: {safe_path}")

        try:
            df = pd.read_xml(safe_path)
            return [df.to_dict(orient='records')]
        except Exception as e:
            logging.debug(f"Default XML parser failed, trying etree: {e}")
            try:
                df = pd.read_xml(safe_path, parser='etree')
                return [df.to_dict(orient='records')]
            except Exception as e2:
                logging.error(f"Error loading XML file: {e2}")
                return []

    def map_and_clean(self, list_of_tables):
        """
        Processes a list of tables (List[List[dict]]) and returns a unified List[dict]
        following the standard simplified CSV format.
        """
        # Backward compatibility: wrap in list if a single table is passed
        if list_of_tables and isinstance(list_of_tables[0], dict):
            list_of_tables = [list_of_tables]

        final_data = []
        generator = Generator()

        for raw_data in list_of_tables:
            if not raw_data:
                continue

            # Identify mapping for this specific table
            first_row = raw_data[0]
            table_mapping = {}
            used_src_cols = set()

            # Explicitly mapped columns from config
            for target, source in self.mapping.items():
                if source in first_row:
                    table_mapping[target] = source
                    used_src_cols.add(source)

            # Fuzzy match for standard columns
            # Priority order as per memory
            priority_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag']
            for target in priority_order:
                if target in table_mapping:
                    continue

                patterns = self.COLUMN_MAPPING.get(target, [])
                for src_col in first_row.keys():
                    if src_col in used_src_cols:
                        continue

                    src_col_lower = str(src_col).lower()
                    if any(p in src_col_lower for p in patterns):
                        table_mapping[target] = src_col
                        used_src_cols.add(src_col)
                        break

            # Process rows of the table
            for row in raw_data:
                new_row = {}
                # Map standard columns
                for target, source in table_mapping.items():
                    new_row[target] = row.get(source)

                # Copy other columns
                for k, v in row.items():
                    if k not in used_src_cols:
                        new_row[k] = v

                # Check for mandatory fields
                if not new_row.get('Name') and not new_row.get('Address'):
                    continue

                # Normalize Address
                addr_raw = str(new_row.get('Address', '')).strip()
                if addr_raw:
                    # Support Address_Length and Address_Start_Bit formats
                    if '_' in addr_raw:
                        parts = addr_raw.split('_')
                        norm_parts = [generator.normalize_address_val(p) for p in parts]
                        new_row['Address'] = '_'.join(norm_parts)
                    else:
                        new_row['Address'] = generator.normalize_address_val(addr_raw)

                # Normalize Type
                if 'Type' in new_row:
                    new_row['Type'] = self.normalize_type(new_row['Type'])

                # Normalize Action
                if 'Action' in new_row:
                    new_row['Action'] = self.normalize_action(new_row['Action'])

                # Default RegisterType
                if not new_row.get('RegisterType'):
                    new_row['RegisterType'] = 'Holding Register'

                final_data.append(new_row)

        return final_data

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

    parser = argparse.ArgumentParser(description='Extract register information from documentation files.')
    parser.add_argument('input_file', help='Path to the source file (PDF, Excel, CSV, XML).')
    parser.add_argument('-o', '--output', help='Path to the output CSV file.')
    parser.add_argument('--mapping', help='JSON file containing column mapping.')
    parser.add_argument('--sheet', help='Excel sheet name to extract from.')
    parser.add_argument('--pages', help='PDF pages to extract from (comma separated, e.g. 1,2,5).')

    args = parser.parse_args()

    mapping = {}
    if args.mapping:
        try:
            safe_mapping_path = ensure_safe_path(args.mapping)
            with open(safe_mapping_path, 'r') as f:
                mapping = json.load(f)
        except Exception as e:
            logging.error(f"Error loading mapping file: {e}")
            sys.exit(1)

    extractor = Extractor(mapping)

    ext = os.path.splitext(args.input_file)[1].lower()
    all_tables = []

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xls']:
        all_tables = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = None
        if args.pages:
            pages = [int(p.strip()) for p in args.pages.split(',')]
        all_tables = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv':
        all_tables = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml':
        all_tables = extractor.extract_from_xml(args.input_file)
    else:
        logging.error(f"Unsupported file extension: {ext}")
        sys.exit(1)

    if not all_tables:
        logging.error("No data extracted.")
        sys.exit(1)

    mapped_data = extractor.map_and_clean(all_tables)

    if not mapped_data:
        logging.error("No data remained after mapping and cleaning.")
        sys.exit(1)

    # Write to CSV
    fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']

    if args.output:
        safe_output_path = ensure_safe_path(args.output)
        f = open(safe_output_path, 'w', newline='', encoding='utf-8')
    else:
        f = sys.stdout

    writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
    writer.writeheader()
    writer.writerows(mapped_data)

    if args.output:
        f.close()
        logging.info(f"Extracted data saved to {args.output}")

if __name__ == "__main__":
    main()
