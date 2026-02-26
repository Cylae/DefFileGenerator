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
        'u16': 'U16',
        'i16': 'I16',
        'u32': 'U32',
        'i32': 'I32',
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
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
        logging.info(f"Extracting from Excel: {filepath}")
        wb = openpyxl.load_workbook(filepath, data_only=True)

        sheets = [sheet_name] if sheet_name else wb.sheetnames
        all_tables = []

        for name in sheets:
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
            target_pages = pdf.pages
            if pages:
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
                        all_tables.append(data)
        return all_tables

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        try:
            with open(filepath, mode='r', encoding='utf-8-sig') as f:
                content = f.read(2048)
                f.seek(0)
                sniffer = csv.Sniffer()
                dialect = sniffer.sniff(content, delimiters=',;\t')
                reader = csv.DictReader(f, dialect=dialect)
                data = list(reader)
                return [data] if data else []
        except Exception as e:
            logging.error(f"Error extracting from CSV: {e}")
            return []

    def extract_from_xml(self, filepath):
        if not HAS_PANDAS:
            logging.error("pandas is required for XML extraction.")
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
        """Processes multiple tables and cleans them."""
        # Handle single table input for backward compatibility
        if tables and isinstance(tables[0], dict):
            tables = [tables]

        all_mapped_data = []
        generator = Generator()

        for raw_data in tables:
            if not raw_data:
                continue

            # Identify columns for this table
            first_row = raw_data[0]
            col_map = {}
            used_src_cols = set()

            # 1. Explicit mapping
            for target, source in self.mapping.items():
                if source in first_row:
                    col_map[target] = source
                    used_src_cols.add(source)

            # 2. Heuristic mapping
            # Priority order
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag', 'Factor', 'ScaleFactor']
            for target in detection_order:
                if target in col_map:
                    continue

                best_match = None
                patterns = self.COLUMN_MAPPING.get(target, [target.lower()])

                # Check exact matches first
                for k in first_row.keys():
                    if k in used_src_cols: continue
                    if k.lower() == target.lower():
                        best_match = k
                        break

                # Then check patterns
                if not best_match:
                    for k in first_row.keys():
                        if k in used_src_cols: continue
                        k_low = k.lower()
                        if any(p in k_low for p in patterns):
                            best_match = k
                            break

                if best_match:
                    col_map[target] = best_match
                    used_src_cols.add(best_match)

            if 'Name' not in col_map and 'Address' not in col_map:
                continue

            for row in raw_data:
                new_row = {}
                for target, source in col_map.items():
                    val = row.get(source)
                    new_row[target] = val if val is not None else ""

                # Skip header-like rows or empty names/addresses
                name_val = str(new_row.get('Name', '')).lower()
                addr_val = str(new_row.get('Address', '')).lower()
                if name_val in self.COLUMN_MAPPING['Name'] or addr_val in self.COLUMN_MAPPING['Address']:
                    continue

                if not new_row.get('Name') and not new_row.get('Address'):
                    continue

                # Normalization
                if new_row.get('Address'):
                    addr = str(new_row['Address']).strip().replace(',', '')
                    # Handle composite addresses like Addr_Len
                    if '_' in addr:
                        parts = addr.split('_')
                        norm_parts = [generator.normalize_address_val(p) for p in parts]
                        new_row['Address'] = '_'.join(norm_parts)
                    else:
                        new_row['Address'] = generator.normalize_address_val(addr)

                if new_row.get('Type'):
                    new_row['Type'] = self.normalize_type(new_row['Type'])

                if not new_row.get('RegisterType'):
                    new_row['RegisterType'] = 'Holding Register'

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
    raw_data = []

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xls']:
        raw_data = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = None
        if args.pages:
            pages = [int(p.strip()) for p in args.pages.split(',')]
        raw_data = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv':
        raw_data = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml':
        raw_data = extractor.extract_from_xml(args.input_file)
    else:
        logging.error(f"Unsupported file extension: {ext}")
        sys.exit(1)

    if not raw_data:
        logging.error("No data extracted.")
        sys.exit(1)

    mapped_data = extractor.map_and_clean(raw_data)

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
