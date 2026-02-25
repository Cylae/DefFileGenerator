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
        'RegisterType': ['register type', 'reg type', 'modbus type', 'info1'],
        'Address': ['address', 'addr', 'offset', 'register', 'reg', 'modbus address'],
        'Name': ['name', 'description', 'parameter', 'variable', 'signal', 'signal name', 'parameter name'],
        'Type': ['data type', 'datatype', 'type', 'format', 'data format'],
        'Unit': ['unit', 'units'],
        'Factor': ['factor', 'multiplier', 'scale', 'ratio', 'scaling'],
        'Action': ['action', 'access', 'access type'],
        'Tag': ['tag'],
        'Offset': ['offset_val'], # To avoid conflict with Address synonym 'offset'
        'ScaleFactor': ['scalefactor']
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
                logging.warning(f"Sheet '{name}' not found in {filepath}")
                continue
            ws = wb[name]
            data = []
            rows = list(ws.rows)
            if not rows:
                continue

            headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]

            for row_idx, row in enumerate(rows[1:], start=2):
                row_data = {}
                for i, cell in enumerate(row):
                    if i < len(headers):
                        row_data[headers[i]] = cell.value
                data.append(row_data)
            if data:
                tables.append(data)
        return tables

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        tables = []
        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                content = f.read(4096)
                f.seek(0)
                sniffer = csv.Sniffer()
                try:
                    dialect = sniffer.sniff(content, delimiters=',;\t')
                except csv.Error:
                    dialect = csv.excel
                    dialect.delimiter = ','

                reader = csv.DictReader(f, dialect=dialect)
                data = list(reader)
                if data:
                    tables.append(data)
        except Exception as e:
            logging.error(f"Error reading CSV {filepath}: {e}")
        return tables

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        logging.info(f"Extracting from PDF: {filepath}")
        tables_data = []
        with pdfplumber.open(filepath) as pdf:
            if pages is None:
                target_pages = pdf.pages
            else:
                target_pages = []
                if isinstance(pages, int):
                    target_pages = [pdf.pages[pages-1]]
                elif isinstance(pages, list):
                    target_pages = [pdf.pages[i-1] for i in pages if 0 < i <= len(pdf.pages)]

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
        return tables_data

    def map_and_clean(self, tables_raw, address_offset=0):
        """Processes a list of tables (list of lists of dicts)."""
        if not tables_raw:
            return []

        # Handle single table input for backward compatibility
        if isinstance(tables_raw, list) and len(tables_raw) > 0 and isinstance(tables_raw[0], dict):
            tables_raw = [tables_raw]

        mapped_data = []
        generator = Generator(address_offset=address_offset)

        for table in tables_raw:
            if not table:
                continue

            first_row = table[0]
            standard_cols_mapping = {}
            assigned_src_cols = set()

            # 1. Explicitly mapped columns from config
            for target, source in self.mapping.items():
                if source in first_row:
                    standard_cols_mapping[target] = source
                    assigned_src_cols.add(source)

            # 2. Fuzzy match for standard columns using COLUMN_MAPPING
            # Priority order for detection
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Factor', 'Action', 'Tag', 'Offset', 'ScaleFactor']
            for target in detection_order:
                if target in standard_cols_mapping:
                    continue

                patterns = self.COLUMN_MAPPING.get(target, [target.lower()])
                for src_col in first_row.keys():
                    if src_col in assigned_src_cols:
                        continue

                    src_col_lower = str(src_col).lower()
                    if any(p in src_col_lower for p in patterns):
                        standard_cols_mapping[target] = src_col
                        assigned_src_cols.add(src_col)
                        break

            for row in table:
                new_row = {}
                for target, source in standard_cols_mapping.items():
                    if source in row:
                        new_row[target] = row[source]

                # Fill in other columns that might not be in standard_cols but are in row
                for k, v in row.items():
                    if k not in assigned_src_cols and k not in new_row:
                        new_row[k] = v

                # Skip rows without a name or address
                if (not new_row.get('Name')) and (not new_row.get('Address')):
                    continue

                # Clean Address using Generator's logic
                if 'Address' in new_row and new_row['Address']:
                    addr = str(new_row['Address']).strip()
                    if '_' in addr:
                        parts = addr.split('_')
                        norm_parts = [generator.normalize_address_val(p) for p in parts]
                        # Apply offset to the first part (base address)
                        try:
                            base_addr = int(norm_parts[0]) - address_offset
                            norm_parts[0] = str(base_addr)
                        except (ValueError, IndexError):
                            pass
                        new_row['Address'] = '_'.join(norm_parts)
                    else:
                        norm_addr = generator.normalize_address_val(addr)
                        try:
                            base_addr = int(norm_addr) - address_offset
                            new_row['Address'] = str(base_addr)
                        except ValueError:
                            new_row['Address'] = norm_addr

                # Clean Type
                if 'Type' in new_row:
                    new_row['Type'] = self.normalize_type(new_row['Type'])

                # Default RegisterType
                if 'RegisterType' not in new_row:
                    new_row['RegisterType'] = 'Holding Register'

                mapped_data.append(new_row)

        return mapped_data

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
    elif ext == '.csv':
        raw_data = extractor.extract_from_csv(args.input_file)
    elif ext == '.pdf':
        pages = None
        if args.pages:
            pages = [int(p.strip()) for p in args.pages.split(',')]
        raw_data = extractor.extract_from_pdf(args.input_file, pages)
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
