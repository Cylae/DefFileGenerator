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

def is_na(val):
    if HAS_PANDAS:
        return pd.isna(val)
    return val is None or val == '' or (isinstance(val, float) and math.isnan(val))

class MockDF:
    def __init__(self, rows, columns):
        self.rows = rows
        self.columns = columns
    def iterrows(self):
        for i, row in enumerate(self.rows):
            yield i, row

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
            'uint64': 'U64',
            'int64': 'I64',
            'float32': 'F32',
            'float': 'F32',
            'f32': 'F32',
            'double': 'F64',
            'f64': 'F64',
            'float64': 'F64',
            'u16': 'U16',
            'i16': 'I16',
            'u32': 'U32',
            'i32': 'I32',
            'u64': 'U64',
            'i64': 'I64',
            'string': 'STRING',
            'bits': 'BITS'
        }

    def normalize_type(self, t):
        if is_na(t):
            return 'U16'
        t_str = str(t).lower().strip()

        # Check explicit mapping first
        for key, val in self.type_mapping.items():
            if key in t_str:
                return val

        # Remove common extra words and spaces
        t_str = t_str.replace('unsigned ', 'u').replace('signed ', 'i').replace(' ', '')

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

    def extract_from_excel(self, filepath, sheet_name=None):
        if HAS_PANDAS:
            try:
                if sheet_name:
                    df = pd.read_excel(filepath, sheet_name=sheet_name)
                    return [df]
                else:
                    excel_file = pd.ExcelFile(filepath)
                    return [excel_file.parse(sheet) for sheet in excel_file.sheet_names]
            except Exception as e:
                logging.error(f"Error loading Excel file via pandas: {e}")
                # Fallback to openpyxl if pandas fails

        if not HAS_OPENPYXL:
            logging.error("pandas or openpyxl is required for Excel extraction.")
            return []

        logging.info(f"Extracting from Excel using openpyxl: {filepath}")
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            results = []
            if sheet_name:
                sheets = [sheet_name] if sheet_name in wb.sheetnames else []
            else:
                sheets = wb.sheetnames

            for name in sheets:
                ws = wb[name]
                rows = list(ws.rows)
                if not rows:
                    continue
                headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]
                data_rows = []
                for row in rows[1:]:
                    row_data = {}
                    for i, cell in enumerate(row):
                        if i < len(headers):
                            row_data[headers[i]] = cell.value
                    data_rows.append(row_data)
                results.append(MockDF(data_rows, headers))
            return results
        except Exception as e:
            logging.error(f"Error loading Excel file via openpyxl: {e}")
            return []

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        logging.info(f"Extracting from PDF: {filepath}")
        dfs = []
        try:
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
                        headers = [str(c).replace('\n', ' ').strip() if c else "" for c in table[0]]
                        data_rows = []
                        for row in table[1:]:
                            row_data = {}
                            for i, cell in enumerate(row):
                                if i < len(headers):
                                    val = str(cell).replace('\n', ' ').strip() if cell else ""
                                    row_data[headers[i]] = val
                            data_rows.append(row_data)
                        dfs.append(MockDF(data_rows, headers))
        except Exception as e:
            logging.error(f"Error loading PDF file: {e}")
        return dfs

    def extract_from_csv(self, filepath):
        if HAS_PANDAS:
            for delimiter in [',', ';', '\t']:
                try:
                    df = pd.read_csv(filepath, sep=delimiter, encoding='utf-8-sig')
                    if len(df.columns) > 1:
                        return [df]
                except Exception:
                    continue

        # Fallback to standard csv module
        for delimiter in [',', ';', '\t']:
             try:
                 with open(filepath, 'r', encoding='utf-8-sig') as f:
                     reader = csv.DictReader(f, delimiter=delimiter)
                     rows = list(reader)
                     if rows and len(reader.fieldnames) > 1:
                         return [MockDF(rows, reader.fieldnames)]
             except Exception:
                 continue
        return []

    def extract_from_xml(self, filepath):
        if not HAS_PANDAS:
            logging.error("pandas is required for XML extraction.")
            return []
        try:
            df = pd.read_xml(filepath)
            return [df]
        except Exception as e:
            logging.debug(f"Default XML parser failed, trying etree: {e}")
            try:
                df = pd.read_xml(filepath, parser='etree')
                return [df]
            except Exception as e2:
                logging.error(f"Error loading XML file: {e2}")
                return []

    def map_and_clean(self, tables):
        all_mapped_data = []
        generator = Generator()

        for df in tables:
            # Clean column names
            columns = [str(c).strip() for c in df.columns]

            # Identify columns
            col_map = {}
            assigned_cols = set()

            # Priority order for detection to avoid misidentification
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Factor', 'ScaleFactor', 'Action', 'Tag']
            for target in detection_order:
                for col in columns:
                    if col in assigned_cols:
                        continue
                    col_lower = str(col).lower()
                    for pattern in self.COLUMN_MAPPING.get(target, []):
                        if pattern in col_lower:
                            col_map[target] = col
                            assigned_cols.add(col)
                            break
                    if target in col_map:
                        break

            if 'Address' not in col_map and 'Name' not in col_map:
                 continue

            # Process rows
            it = df.iterrows() if hasattr(df, 'iterrows') else enumerate(df.rows)
            for _, row in it:
                new_row = {}

                addr_raw = row.get(col_map.get('Address')) if 'Address' in col_map else None
                name_raw = row.get(col_map.get('Name')) if 'Name' in col_map else None

                if is_na(addr_raw) and is_na(name_raw):
                    continue

                # Skip header-like rows
                if str(addr_raw).lower() in self.COLUMN_MAPPING['Address'] or str(name_raw).lower() in self.COLUMN_MAPPING['Name']:
                    continue

                # Basic extraction using identified columns
                new_row['Name'] = str(name_raw) if not is_na(name_raw) else ""
                new_row['Address'] = str(addr_raw) if not is_na(addr_raw) else ""
                new_row['Tag'] = str(row.get(col_map.get('Tag'))) if 'Tag' in col_map and not is_na(row.get(col_map.get('Tag'))) else ""
                new_row['RegisterType'] = str(row.get(col_map.get('RegisterType'))) if 'RegisterType' in col_map and not is_na(row.get(col_map.get('RegisterType'))) else 'Holding Register'
                new_row['Type'] = self.normalize_type(row.get(col_map.get('Type'))) if 'Type' in col_map else 'U16'
                new_row['Unit'] = str(row.get(col_map.get('Unit'))) if 'Unit' in col_map and not is_na(row.get(col_map.get('Unit'))) else ""

                # Handle Factor/Scale (Factor in target maps to Scale synonyms)
                factor_raw = row.get(col_map.get('Factor')) if 'Factor' in col_map else '1'
                if is_na(factor_raw): factor_raw = '1'
                # Support fractions like 1/10
                factor_str = str(factor_raw)
                if '/' in factor_str:
                    try:
                        parts = factor_str.split('/')
                        factor_str = str(float(parts[0]) / float(parts[1]))
                    except (ValueError, ZeroDivisionError):
                        factor_str = '1'
                new_row['Factor'] = factor_str

                new_row['ScaleFactor'] = str(row.get(col_map.get('ScaleFactor'))) if 'ScaleFactor' in col_map and not is_na(row.get(col_map.get('ScaleFactor'))) else '0'
                new_row['Action'] = str(row.get(col_map.get('Action'))) if 'Action' in col_map and not is_na(row.get(col_map.get('Action'))) else '1'
                new_row['Offset'] = '0' # Default

                # Delegate final address normalization to Generator
                if new_row['Address']:
                    addr = new_row['Address'].strip()
                    # Handle formats like 40,001
                    if ',' in addr and '.' not in addr:
                        addr = addr.replace(',', '')

                    # Extract the address part(s) - handle 30001_10 or 0x10_0_1
                    match = re.search(r'(0x[0-9A-Fa-f]+|\d+(_\d+)*|[0-9A-Fa-f]+h(_\d+)*)', addr)
                    if match:
                        addr = match.group(1)

                    if '_' in addr:
                        parts = addr.split('_')
                        norm_parts = [generator.normalize_address_val(p) for p in parts]
                        new_row['Address'] = '_'.join(norm_parts)
                    else:
                        new_row['Address'] = generator.normalize_address_val(addr)

                if not new_row['Name'] and not new_row['Address']:
                    continue

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
