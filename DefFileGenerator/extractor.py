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

        if t_str in self.TYPE_MAPPING:
            return self.TYPE_MAPPING[t_str]

        # Check for patterns like Uint16, Int32, uint16, int32
        match = re.match(r'^(u|i|uint|int)(\d+)$', t_str)
        if match:
            raw_prefix = match.group(1).lower()
            prefix = 'U' if raw_prefix.startswith('u') else 'I'
            bits = match.group(2)
            return f"{prefix}{bits}"

        return t.upper()

    def extract_from_excel(self, filepath, sheet_name=None):
        ext = os.path.splitext(filepath)[1].lower()
        if ext == '.xls':
            if not HAS_PANDAS:
                logging.error("pandas and xlrd are required for .xls extraction.")
                return []
            logging.info(f"Extracting from Excel (.xls): {filepath}")
            try:
                if sheet_name:
                    df = pd.read_excel(filepath, sheet_name=sheet_name)
                    return [df.to_dict(orient='records')]
                else:
                    xl = pd.ExcelFile(filepath)
                    return [xl.parse(sheet).to_dict(orient='records') for sheet in xl.sheet_names]
            except Exception as e:
                logging.error(f"Error extracting from .xls: {e}")
                return []

        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []

        logging.info(f"Extracting from Excel: {filepath}")
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            tables = []
            sheets_to_process = [sheet_name] if sheet_name else wb.sheetnames

            for name in sheets_to_process:
                if name not in wb.sheetnames:
                    logging.warning(f"Sheet '{name}' not found in {filepath}")
                    continue
                ws = wb[name]
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
                    tables.append(data)
            return tables
        except Exception as e:
            logging.error(f"Error extracting from Excel: {e}")
            return []

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        logging.info(f"Extracting from PDF: {filepath}")
        tables_data = []
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

                        # Clean headers: remove newlines
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
                            tables_data.append(data)
        except Exception as e:
            logging.error(f"Error extracting from PDF: {e}")
        return tables_data

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        # Use utf-8-sig to handle potential BOM
        for delimiter in [',', ';', '\t']:
            try:
                with open(filepath, 'r', encoding='utf-8-sig') as f:
                    reader = csv.DictReader(f, delimiter=delimiter)
                    rows = list(reader)
                    if rows and len(reader.fieldnames) > 1:
                        # Clean fieldnames
                        cleaned_rows = []
                        for row in rows:
                            cleaned_row = {fn.strip() if fn else f"Col{i}": v for i, (fn, v) in enumerate(row.items())}
                            cleaned_rows.append(cleaned_row)
                        return [cleaned_rows]
            except Exception:
                continue
        return []

    def extract_from_xml(self, filepath):
        if not HAS_PANDAS:
            logging.error("pandas and lxml are required for XML extraction.")
            return []
        logging.info(f"Extracting from XML: {filepath}")
        try:
            df = pd.read_xml(filepath)
            return [df.to_dict(orient='records')]
        except Exception as e:
            logging.debug(f"Default XML parser failed, trying etree: {e}")
            try:
                df = pd.read_xml(filepath, parser='etree')
                return [df.to_dict(orient='records')]
            except Exception as e2:
                logging.error(f"Error loading XML file: {e2}")
                return []

    def map_and_clean(self, tables):
        if not tables:
            return []

        # Handle backward compatibility: if a single list of dicts is passed instead of list of tables
        if isinstance(tables, list) and len(tables) > 0 and isinstance(tables[0], dict):
            tables = [tables]

        all_mapped_data = []
        generator = Generator()

        for raw_data in tables:
            if not raw_data:
                continue

            # Identify standard columns for this table
            first_row = raw_data[0]
            standard_cols_mapping = {}
            assigned_keys = set()

            # Explicitly mapped columns from config
            for target, source in self.mapping.items():
                if source in first_row:
                    standard_cols_mapping[target] = source
                    assigned_keys.add(source)

            # Fuzzy match for standard columns if not explicitly mapped
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Scale', 'Action']
            for target in detection_order:
                if target not in standard_cols_mapping:
                    for k in first_row.keys():
                        if k in assigned_keys:
                            continue
                        k_lower = str(k).lower()
                        for pattern in self.COLUMN_MAPPING[target]:
                            if pattern in k_lower:
                                standard_cols_mapping[target] = k
                                assigned_keys.add(k)
                                break
                        if target in standard_cols_mapping:
                            break

            if 'Address' not in standard_cols_mapping and 'Name' not in standard_cols_mapping:
                 continue

            for row in raw_data:
                new_row = {}
                for target, source in standard_cols_mapping.items():
                    if source in row:
                        new_row[target] = row[source]

                # Fill in other columns that might not be in standard_cols but are in row
                for k, v in row.items():
                    if k not in assigned_keys and k not in new_row:
                        new_row[k] = v

                # Skip header-like rows
                addr_raw = new_row.get('Address')
                name_raw = new_row.get('Name')
                if addr_raw is None and name_raw is None:
                    continue

                # Check if this row is just repeating headers
                is_header = False
                if addr_raw and str(addr_raw).lower() in self.COLUMN_MAPPING['Address']:
                    is_header = True
                if name_raw and str(name_raw).lower() in self.COLUMN_MAPPING['Name']:
                    is_header = True
                if is_header:
                    continue

                # Clean Address using Generator's logic
                if 'Address' in new_row and new_row['Address'] is not None:
                    addr = str(new_row['Address']).strip()
                    # Handle commas as thousands separators
                    if ',' in addr and '.' not in addr:
                        addr = addr.replace(',', '')

                    if '_' in addr:
                        parts = addr.split('_')
                        norm_parts = [generator.normalize_address_val(p) for p in parts]
                        new_row['Address'] = '_'.join(norm_parts)
                    else:
                        new_row['Address'] = generator.normalize_address_val(addr)

                # Clean Type
                if 'Type' in new_row:
                    new_row['Type'] = self.normalize_type(new_row['Type'])

                # Clean Action
                if 'Action' in new_row:
                    a = str(new_row['Action']).upper().strip()
                    if a == 'R':
                        new_row['Action'] = '4'
                    elif a == 'RW' or a == 'W':
                        new_row['Action'] = '1'

                # Clean Scale/Factor
                if 'Scale' in new_row and new_row['Scale'] is not None:
                    scale = str(new_row['Scale'])
                    if '/' in scale:
                        try:
                            parts = scale.split('/')
                            new_row['Factor'] = str(float(parts[0]) / float(parts[1]))
                        except (ValueError, ZeroDivisionError):
                            new_row['Factor'] = '1'
                    else:
                        new_row['Factor'] = scale

                    # Remove Scale key to avoid confusion
                    new_row.pop('Scale', None)

                # Ensure mandatory fields for def_gen
                if not new_row.get('Name') and not new_row.get('Address'):
                    continue

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
