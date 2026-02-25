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
        if t is None or (isinstance(t, float) and math.isnan(t)) or t == '':
            return 'U16'
        t_str = str(t).lower().strip()
        # Remove common extra words and spaces
        t_str = t_str.replace('unsigned ', 'u').replace('signed ', 'i').replace(' ', '')

        if t_str in self.TYPE_MAPPING:
            return self.TYPE_MAPPING[t_str]

        # Check for partial matches in TYPE_MAPPING
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

        return str(t).upper().replace(' ', '')

    def normalize_action(self, action):
        if action is None or (isinstance(action, float) and math.isnan(action)) or action == '':
            return '1'
        a = str(action).upper().strip()
        if a in ['R', 'READ']:
            return '4'
        if a in ['RW', 'W', 'WRITE']:
            return '1'
        return a

    def extract_from_excel(self, filepath, sheet_name=None):
        logging.info(f"Extracting from Excel: {filepath}")
        all_tables = []
        if HAS_PANDAS:
            try:
                if sheet_name:
                    df = pd.read_excel(filepath, sheet_name=sheet_name)
                    all_tables.append(df.to_dict(orient='records'))
                else:
                    excel_file = pd.ExcelFile(filepath)
                    for sheet in excel_file.sheet_names:
                        df = excel_file.parse(sheet)
                        if not df.empty:
                            all_tables.append(df.to_dict(orient='records'))
            except Exception as e:
                logging.error(f"Error extracting from Excel using pandas: {e}")
        elif HAS_OPENPYXL:
            try:
                wb = openpyxl.load_workbook(filepath, data_only=True)
                sheets = [wb[sheet_name]] if sheet_name else wb.worksheets
                for ws in sheets:
                    data = []
                    rows = list(ws.rows)
                    if not rows: continue
                    headers = [str(cell.value).strip() if cell.value is not None else f"Col{i}" for i, cell in enumerate(rows[0])]
                    for row in rows[1:]:
                        row_data = {}
                        for i, cell in enumerate(row):
                            if i < len(headers):
                                row_data[headers[i]] = cell.value
                        if any(v is not None for v in row_data.values()):
                            data.append(row_data)
                    if data:
                        all_tables.append(data)
            except Exception as e:
                logging.error(f"Error extracting from Excel using openpyxl: {e}")
        else:
            logging.error("pandas or openpyxl is required for Excel extraction.")
        return all_tables

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        logging.info(f"Extracting from PDF: {filepath}")
        all_tables = []
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
                        headers = [str(c).replace('\n', ' ').strip() if c else f"Col{i}" for i, c in enumerate(table[0])]
                        data = []
                        for row in table[1:]:
                            row_data = {}
                            for i, cell in enumerate(row):
                                if i < len(headers):
                                    val = str(cell).replace('\n', ' ').strip() if cell else ""
                                    row_data[headers[i]] = val
                            if any(v for v in row_data.values()):
                                data.append(row_data)
                        if data:
                            all_tables.append(data)
        except Exception as e:
            logging.error(f"Error extracting from PDF: {e}")
        return all_tables

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        if HAS_PANDAS:
            for delimiter in [',', ';', '\t']:
                try:
                    df = pd.read_csv(filepath, sep=delimiter, encoding='utf-8-sig')
                    if len(df.columns) > 1:
                        return [df.to_dict(orient='records')]
                except Exception:
                    continue
        else:
            for delimiter in [',', ';', '\t']:
                try:
                    with open(filepath, 'r', encoding='utf-8-sig') as f:
                        reader = csv.DictReader(f, delimiter=delimiter)
                        rows = list(reader)
                        if rows and len(reader.fieldnames) > 1:
                            return [rows]
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
                logging.debug(f"Pandas read_xml failed, trying etree: {e}")
                try:
                    df = pd.read_xml(filepath, parser='etree')
                    return [df.to_dict(orient='records')]
                except Exception as e2:
                    logging.error(f"Error extracting from XML: {e2}")
        else:
            logging.error("pandas and lxml are required for XML extraction.")
        return []

    def map_and_clean(self, tables):
        """Processes list of tables (each a list of dicts) into a unified format."""
        if not tables:
            return []

        # Support both single table (backward compatibility) and list of tables
        if isinstance(tables, list) and len(tables) > 0 and isinstance(tables[0], dict):
            tables = [tables]

        all_mapped_data = []
        generator = Generator()

        for table in tables:
            if not table:
                continue

            first_row = table[0]
            col_map = {}
            used_src_cols = set()

            # Explicitly mapped columns from config
            for target, source in self.mapping.items():
                if source in first_row:
                    col_map[target] = source
                    used_src_cols.add(source)

            # Fuzzy match for standard columns
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Factor', 'ScaleFactor', 'Action', 'Tag']
            for target in detection_order:
                if target in col_map:
                    continue
                patterns = self.COLUMN_MAPPING.get(target, [])
                for src_col in first_row.keys():
                    if src_col in used_src_cols:
                        continue
                    src_lower = str(src_col).lower()
                    if any(p in src_lower for p in patterns):
                        col_map[target] = src_col
                        used_src_cols.add(src_col)
                        break

            if 'Address' not in col_map and 'Name' not in col_map:
                continue

            for row in table:
                new_row = {}
                for target, source in col_map.items():
                    new_row[target] = row.get(source)

                # Fill remaining columns
                for k, v in row.items():
                    if k not in used_src_cols and k not in new_row:
                        new_row[k] = v

                # Normalization
                if new_row.get('Name') is None and new_row.get('Address') is None:
                    continue

                if 'Address' in new_row:
                    addr_str = str(new_row['Address']).strip() if new_row['Address'] is not None else ""
                    if ',' in addr_str and '.' not in addr_str:
                        addr_str = addr_str.replace(',', '')

                    # Support Address_Length and Address_Start_Bit
                    if '_' in addr_str:
                        parts = addr_str.split('_')
                        norm_parts = [generator.normalize_address_val(p) for p in parts]
                        new_row['Address'] = '_'.join(norm_parts)
                    else:
                        # Extract address using regex if it's messy
                        match = re.search(r'(0x[0-9A-Fa-f]+|[0-9A-Fa-f]+h|\d+)', addr_str, re.IGNORECASE)
                        if match:
                            new_row['Address'] = generator.normalize_address_val(match.group(1))
                        else:
                            new_row['Address'] = generator.normalize_address_val(addr_str)

                if 'Type' in new_row:
                    new_row['Type'] = self.normalize_type(new_row['Type'])

                if 'Action' in new_row:
                    new_row['Action'] = self.normalize_action(new_row['Action'])

                if 'Factor' in new_row and new_row['Factor']:
                    # Support fractions like 1/10
                    f_str = str(new_row['Factor'])
                    if '/' in f_str:
                        try:
                            parts = f_str.split('/')
                            new_row['Factor'] = str(float(parts[0]) / float(parts[1]))
                        except (ValueError, ZeroDivisionError):
                            pass

                if not new_row.get('Name'):
                    if new_row.get('Address'):
                        new_row['Name'] = f"Register {new_row['Address']}"
                    else:
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
    parser.add_argument('--pages', help='PDF pages to extract from (comma separated).')

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
