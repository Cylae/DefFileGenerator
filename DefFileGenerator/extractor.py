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
        'float32': 'F32',
        'float': 'F32',
        'f32': 'F32',
        'double': 'F64',
        'f64': 'F64',
        'float64': 'F64',
    }

    def __init__(self, mapping=None):
        self.mapping = mapping or {}

    def normalize_type(self, t):
        if t is None or (HAS_PANDAS and pd.isna(t)):
            return 'U16'
        t_str = str(t).lower().strip()

        # Apply type mapping synonyms
        for key, val in self.TYPE_MAPPING.items():
            if key in t_str:
                return val

        # Clean up common extra words and spaces
        t_str = t_str.replace('unsigned', 'u').replace('signed', 'i').replace(' ', '')

        # Check for patterns like uint16, int32
        match = re.match(r'^(u|i|uint|int)(\d+)$', t_str)
        if match:
            raw_prefix = match.group(1).lower()
            prefix = 'U' if raw_prefix.startswith('u') else 'I'
            bits = match.group(2)
            return f"{prefix}{bits}"

        # Final cleanup for Webdyn format
        t_str = re.sub(r'[^a-z0-9_]+', '', t_str)
        return t_str.upper() if t_str else 'U16'

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
        logging.info(f"Extracting from Excel: {filepath}")
        wb = openpyxl.load_workbook(filepath, data_only=True)

        all_tables = []
        sheets_to_process = [sheet_name] if sheet_name else wb.sheetnames

        for name in sheets_to_process:
            if name not in wb.sheetnames:
                logging.warning(f"Sheet '{name}' not found.")
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
                for p in pages:
                    if 0 < p <= len(pdf.pages):
                        target_pages.append(pdf.pages[p-1])

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
        logging.info(f"Extracting from CSV: {filepath}")
        for delimiter in [',', ';', '\t']:
            try:
                with open(filepath, 'r', encoding='utf-8-sig') as f:
                    # Check if delimiter exists in first 4KB
                    content = f.read(4096)
                    f.seek(0)
                    if delimiter in content:
                        reader = csv.DictReader(f, delimiter=delimiter)
                        rows = list(reader)
                        if rows and len(reader.fieldnames) > 1:
                            return [rows]
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
            logging.error(f"Error extracting from XML: {e}")
            return []

    def map_and_clean(self, raw_tables):
        """
        Processes a list of tables. raw_tables can be List[List[dict]] or List[dict].
        """
        if not raw_tables:
            return []

        # Handle backward compatibility if single table passed
        if isinstance(raw_tables[0], dict):
            raw_tables = [raw_tables]

        all_mapped_data = []
        generator = Generator()

        for table in raw_tables:
            if not table:
                continue

            # Identify columns for this table
            first_row = table[0]
            col_map = {}
            assigned_src_cols = set()

            # Priority match using COLUMN_MAPPING
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Factor', 'ScaleFactor', 'Tag', 'Action']

            # First, use explicit mapping
            for target, source in self.mapping.items():
                if source in first_row:
                    col_map[target] = source
                    assigned_src_cols.add(source)

            # Then, use heuristics
            for target in detection_order:
                if target in col_map:
                    continue
                for src_col in first_row.keys():
                    if src_col in assigned_src_cols:
                        continue
                    src_col_lower = str(src_col).lower()
                    for pattern in self.COLUMN_MAPPING.get(target, []):
                        if pattern in src_col_lower:
                            col_map[target] = src_col
                            assigned_src_cols.add(src_col)
                            break
                    if target in col_map:
                        break

            if 'Address' not in col_map and 'Name' not in col_map:
                continue

            for row in table:
                addr_raw = row.get(col_map.get('Address')) if 'Address' in col_map else None
                name_raw = row.get(col_map.get('Name')) if 'Name' in col_map else None

                if (addr_raw is None or addr_raw == '') and (name_raw is None or name_raw == ''):
                    continue

                # Skip header-like rows
                if str(addr_raw).lower() in self.COLUMN_MAPPING['Address'] or str(name_raw).lower() in self.COLUMN_MAPPING['Name']:
                    continue

                new_row = {
                    'Name': str(name_raw) if name_raw is not None else '',
                    'Address': str(addr_raw) if addr_raw is not None else '',
                    'Type': self.normalize_type(row.get(col_map.get('Type'))) if 'Type' in col_map else 'U16',
                    'Unit': str(row.get(col_map.get('Unit'))) if 'Unit' in col_map and row.get(col_map.get('Unit')) is not None else '',
                    'RegisterType': str(row.get(col_map.get('RegisterType'))) if 'RegisterType' in col_map else 'Holding Register',
                    'Factor': str(row.get(col_map.get('Factor'))) if 'Factor' in col_map else '1',
                    'ScaleFactor': str(row.get(col_map.get('ScaleFactor'))) if 'ScaleFactor' in col_map else '0',
                    'Offset': '0', # Default
                    'Action': str(row.get(col_map.get('Action'))) if 'Action' in col_map else '1',
                    'Tag': str(row.get(col_map.get('Tag'))) if 'Tag' in col_map else ''
                }

                # Normalize Address (convert any hex parts to decimal and remove commas)
                if new_row['Address']:
                    addr = new_row['Address'].strip().replace(',', '')
                    if '_' in addr:
                        parts = addr.split('_')
                        norm_parts = [generator.normalize_address_val(p) for p in parts]
                        new_row['Address'] = '_'.join(norm_parts)
                    else:
                        new_row['Address'] = generator.normalize_address_val(addr)

                # Action normalization (synonyms handled by Generator.process_rows,
                # but we can do a quick pass here too if we want, but it's better to let Generator do it)

                # Cleanup factor if it's "1/10"
                if '/' in new_row['Factor']:
                    try:
                        parts = new_row['Factor'].split('/')
                        new_row['Factor'] = str(float(parts[0]) / float(parts[1]))
                    except (ValueError, ZeroDivisionError):
                        new_row['Factor'] = '1'

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
    raw_tables = []

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xls']:
        raw_tables = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = None
        if args.pages:
            pages = [int(p.strip()) for p in args.pages.split(',')]
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
