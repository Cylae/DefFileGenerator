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
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

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
        'Scale': ['scale', 'factor', 'multiplier', 'ratio'],
        'Action': ['action', 'access'],
        'Tag': ['tag']
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

        # Clean up common characters like () or space
        t_str = re.sub(r'[^a-z0-9_]+', '', t_str)
        return t_str.upper() if t_str else 'U16'

    def normalize_action(self, action):
        if action is None or (isinstance(action, float) and math.isnan(action)) or str(action).strip() == '':
            return '1'
        a = str(action).upper().strip()
        if a == 'R' or 'READ' in a and 'WRITE' not in a:
            return '4'
        if a == 'RW' or a == 'W' or 'WRITE' in a:
            return '1'
        return a

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_PANDAS or not HAS_OPENPYXL:
            logging.error("pandas and openpyxl are required for Excel extraction.")
            return []
        logging.info(f"Extracting from Excel: {filepath}")
        try:
            if sheet_name:
                df_dict = {sheet_name: pd.read_excel(filepath, sheet_name=sheet_name)}
            else:
                df_dict = pd.read_excel(filepath, sheet_name=None)

            tables = []
            for name, df in df_dict.items():
                if not df.empty:
                    # Convert NaN to None for consistent processing
                    tables.append(df.where(pd.notnull(df), None).to_dict(orient='records'))
            return tables
        except Exception as e:
            logging.error(f"Error loading Excel file: {e}")
            return []

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        logging.info(f"Extracting from PDF: {filepath}")
        tables = []
        try:
            with pdfplumber.open(filepath) as pdf:
                if pages is None:
                    target_pages = pdf.pages
                else:
                    target_pages = [pdf.pages[i-1] for i in pages if 0 < i <= len(pdf.pages)]

                for page in target_pages:
                    extracted_tables = page.extract_tables()
                    for table in extracted_tables:
                        if not table or len(table) < 2:
                            continue

                        # Clean headers: remove newlines
                        headers = [str(c).replace('\n', ' ').strip() if c else f"Col{i}" for i, c in enumerate(table[0])]

                        rows = []
                        for row_data in table[1:]:
                            row_dict = {}
                            for i, cell in enumerate(row_data):
                                if i < len(headers):
                                    val = str(cell).replace('\n', ' ').strip() if cell is not None else None
                                    row_dict[headers[i]] = val
                            rows.append(row_dict)
                        tables.append(rows)
            return tables
        except Exception as e:
            logging.error(f"Error loading PDF file: {e}")
            return []

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        for delimiter in [',', ';', '\t']:
            try:
                if HAS_PANDAS:
                    df = pd.read_csv(filepath, sep=delimiter, encoding='utf-8-sig')
                    if len(df.columns) > 1:
                        return [df.where(pd.notnull(df), None).to_dict(orient='records')]
                else:
                    with open(filepath, 'r', encoding='utf-8-sig') as f:
                        reader = csv.DictReader(f, delimiter=delimiter)
                        rows = list(reader)
                        if rows and len(reader.fieldnames) > 1:
                            return [rows]
            except Exception:
                continue
        return []

    def extract_from_xml(self, filepath):
        if not HAS_PANDAS:
            logging.error("pandas is required for XML extraction.")
            return []
        logging.info(f"Extracting from XML: {filepath}")
        try:
            df = pd.read_xml(filepath)
            return [df.where(pd.notnull(df), None).to_dict(orient='records')]
        except Exception as e:
            logging.error(f"Error loading XML file: {e}")
            return []

    def map_and_clean(self, tables):
        if not tables:
            return []

        # Support both List[dict] and List[List[dict]] for backward compatibility
        if isinstance(tables, list) and len(tables) > 0 and isinstance(tables[0], dict):
            tables = [tables]

        all_mapped_data = []
        generator = Generator()

        for table in tables:
            if not table:
                continue

            # Identify columns for this table
            first_row = table[0]
            col_map = {}
            assigned_src_cols = set()

            # Explicitly mapped columns from config
            for target, source in self.mapping.items():
                if source in first_row:
                    col_map[target] = source
                    assigned_src_cols.add(source)

            # Priority-based fuzzy matching
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Scale', 'Action', 'Tag']
            for target in detection_order:
                if target in col_map:
                    continue
                for col in first_row.keys():
                    if col in assigned_src_cols:
                        continue
                    col_lower = str(col).lower()
                    for pattern in self.COLUMN_MAPPING.get(target, []):
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

                # Skip header rows if they got included
                if str(addr_raw).lower() in self.COLUMN_MAPPING['Address'] or str(name_raw).lower() in self.COLUMN_MAPPING['Name']:
                    continue

                new_row = {}
                # Map standard columns
                new_row['Name'] = str(name_raw) if name_raw is not None else ""
                new_row['Address'] = str(addr_raw) if addr_raw is not None else ""
                new_row['Type'] = self.normalize_type(row.get(col_map.get('Type'))) if 'Type' in col_map else 'U16'
                new_row['Unit'] = str(row.get(col_map.get('Unit'))) if 'Unit' in col_map and row.get(col_map.get('Unit')) is not None else ""

                # Action normalization
                new_row['Action'] = self.normalize_action(row.get(col_map.get('Action'))) if 'Action' in col_map else '1'

                # RegisterType
                rt = row.get(col_map.get('RegisterType')) if 'RegisterType' in col_map else 'Holding Register'
                new_row['RegisterType'] = str(rt) if rt is not None else 'Holding Register'

                # Scale/Factor
                scale_raw = row.get(col_map.get('Scale')) if 'Scale' in col_map else '1'
                scale = str(scale_raw) if scale_raw is not None else '1'
                if '/' in scale:
                    try:
                        parts = scale.split('/')
                        scale = str(float(parts[0]) / float(parts[1]))
                    except (ValueError, ZeroDivisionError):
                        scale = '1'
                new_row['Factor'] = scale
                new_row['Offset'] = '0'
                new_row['Tag'] = str(row.get(col_map.get('Tag'))) if 'Tag' in col_map and row.get(col_map.get('Tag')) is not None else ''

                # Fill in any other columns
                for k, v in row.items():
                    if k not in assigned_src_cols and k not in new_row:
                        new_row[k] = v

                # Final Address normalization delegating to Generator
                if new_row['Address']:
                    addr = new_row['Address'].strip()
                    if '_' in addr:
                        parts = addr.split('_')
                        norm_parts = [generator.normalize_address_val(p) for p in parts]
                        new_row['Address'] = '_'.join(norm_parts)
                    else:
                        new_row['Address'] = generator.normalize_address_val(addr)

                if not new_row['Name'] and not new_row['Address']:
                    continue

                if not new_row['Name']:
                    new_row['Name'] = f"Register {new_row['Address']}"

                all_mapped_data.append(new_row)

        return all_mapped_data

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

    parser = argparse.ArgumentParser(description='Extract register information from PDF/Excel/CSV/XML files.')
    parser.add_argument('input_file', help='Path to the source file.')
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
    tables = []

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xls']:
        tables = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = None
        if args.pages:
            pages = [int(p.strip()) for p in args.pages.split(',')]
        tables = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv':
        tables = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml':
        tables = extractor.extract_from_xml(args.input_file)
    else:
        logging.error(f"Unsupported file extension: {ext}")
        sys.exit(1)

    if not tables:
        logging.error("No data extracted.")
        sys.exit(1)

    mapped_data = extractor.map_and_clean(tables)

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
