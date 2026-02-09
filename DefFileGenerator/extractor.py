#!/usr/bin/env python3
import argparse
import csv
import json
import logging
import os
import re
import sys
import openpyxl
import pdfplumber

class Extractor:
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

    def parse_address(self, addr):
        if not addr:
            return ""
        addr_str = str(addr).strip()
        # Handle Hex: 0x1234 or 1234H
        if addr_str.lower().startswith('0x'):
            try:
                return str(int(addr_str, 16))
            except ValueError:
                return addr_str
        if addr_str.lower().endswith('h'):
            try:
                return str(int(addr_str[:-1], 16))
            except ValueError:
                return addr_str

        # Handle composite address formats (e.g., 30001_10 or 30001_0_1)
        if '_' in addr_str:
            parts = addr_str.split('_')
            if all(p.isdigit() for p in parts):
                return addr_str

        return addr_str

    def extract_from_excel(self, filepath, sheet_name=None):
        logging.info(f"Extracting from Excel: {filepath}")
        wb = openpyxl.load_workbook(filepath, data_only=True)
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                logging.error(f"Sheet '{sheet_name}' not found in {filepath}")
                return []
            ws = wb[sheet_name]
        else:
            ws = wb.active

        data = []
        rows = list(ws.rows)
        if not rows:
            return []

        headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]

        for row_idx, row in enumerate(rows[1:], start=2):
            row_data = {}
            for i, cell in enumerate(row):
                if i < len(headers):
                    row_data[headers[i]] = cell.value
            data.append(row_data)
        return data

    def extract_from_pdf(self, filepath, pages=None):
        logging.info(f"Extracting from PDF: {filepath}")
        data = []
        with pdfplumber.open(filepath) as pdf:
            if pages is None:
                target_pages = pdf.pages
            else:
                target_pages = []
                # Simple page selection logic
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

                    for row in table[1:]:
                        row_data = {}
                        for i, cell in enumerate(row):
                            if i < len(headers):
                                val = str(cell).replace('\n', ' ').strip() if cell else ""
                                row_data[headers[i]] = val
                        data.append(row_data)
        return data

    def map_and_clean(self, raw_data):
        mapped_data = []
        # Mapping: target_col -> source_col
        # Default target cols: Name, Tag, RegisterType, Address, Type, Factor, Offset, Unit, Action, ScaleFactor

        for row in raw_data:
            new_row = {}
            # Apply mapping from config
            for target, source in self.mapping.items():
                if source in row:
                    new_row[target] = row[source]

            # Fuzzy match for standard columns if not explicitly mapped
            # Priority order: RegisterType > Address > Name > Type > Unit > Tag
            standard_cols = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Tag']
            used_src_cols = set(self.mapping.values())

            for col in standard_cols:
                if col not in new_row:
                    # Try to find a match in row keys
                    for k in row.keys():
                        if k in used_src_cols:
                            continue
                        if k.lower() == col.lower() or col.lower() in k.lower():
                            new_row[col] = row[k]
                            used_src_cols.add(k)
                            break

            # Clean Address
            if 'Address' in new_row:
                new_row['Address'] = self.parse_address(new_row['Address'])

            # Clean Type
            if 'Type' in new_row:
                new_row['Type'] = self.normalize_type(new_row['Type'])

            # Ensure mandatory fields for def_gen
            if 'Name' not in new_row or not new_row['Name']:
                continue # Skip rows without a name

            if 'Tag' not in new_row or not new_row['Tag']:
                # Generate tag from name
                name = str(new_row['Name'])
                tag = re.sub(r'[^a-z0-9_]', '', name.lower().replace(' ', '_'))
                new_row['Tag'] = tag

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

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw_data = extractor.extract_from_excel(args.input_file, args.sheet)
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
