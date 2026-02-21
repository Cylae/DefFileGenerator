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

        all_tables = []
        sheets = [sheet_name] if sheet_name else wb.sheetnames

        for name in sheets:
            if name not in wb.sheetnames:
                logging.error(f"Sheet '{name}' not found in {filepath}")
                continue
            ws = wb[name]
            rows = list(ws.rows)
            if not rows:
                continue

            headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]
            table_data = []
            for row in rows[1:]:
                row_data = {}
                for i, cell in enumerate(row):
                    if i < len(headers):
                        row_data[headers[i]] = cell.value
                table_data.append(row_data)
            all_tables.append(table_data)
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
                    if 1 <= p <= len(pdf.pages):
                        target_pages.append(pdf.pages[p-1])

            for page in target_pages:
                tables = page.extract_tables()
                for table in tables:
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
                    all_tables.append(table_data)
        return all_tables

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        # Try different delimiters
        for delimiter in [',', ';', '\t']:
            try:
                with open(filepath, mode='r', encoding='utf-8-sig') as f:
                    # Sniffing can be unreliable, so we manually check for delimiter presence in first line
                    first_line = f.readline()
                    f.seek(0)
                    if delimiter not in first_line and delimiter != ',':
                        continue

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

    def map_and_clean(self, raw_data):
        if not raw_data:
            return []

        # Handle both single table (list of dicts) and list of tables (list of lists of dicts)
        if isinstance(raw_data, list) and len(raw_data) > 0 and isinstance(raw_data[0], list):
            all_mapped = []
            for table in raw_data:
                all_mapped.extend(self._process_single_table(table))
            return all_mapped
        else:
            return self._process_single_table(raw_data)

    def _process_single_table(self, table_data):
        if not table_data:
            return []

        mapped_data = []
        first_row = table_data[0]
        standard_cols_mapping = {}
        assigned_keys = set()

        # Explicitly mapped columns from config
        for target, source in self.mapping.items():
            if source in first_row:
                standard_cols_mapping[target] = source
                assigned_keys.add(source)

        # Fuzzy match for standard columns
        standard_cols = ['RegisterType', 'Name', 'Address', 'Type', 'Unit', 'Tag']
        for target in standard_cols:
            if target not in standard_cols_mapping:
                for k in first_row.keys():
                    if k in assigned_keys:
                        continue
                    if k.lower() == target.lower() or target.lower() in k.lower():
                        standard_cols_mapping[target] = k
                        assigned_keys.add(k)
                        break

        generator = Generator()
        for row in table_data:
            new_row = {}
            for target, source in standard_cols_mapping.items():
                if source in row:
                    new_row[target] = row[source]

            for k, v in row.items():
                if k not in assigned_keys and k not in new_row:
                    new_row[k] = v

            if 'Address' in new_row and new_row['Address']:
                addr = str(new_row['Address']).strip()
                if '_' in addr:
                    parts = addr.split('_')
                    norm_parts = [generator.normalize_address_val(p) for p in parts]
                    new_row['Address'] = '_'.join(norm_parts)
                else:
                    new_row['Address'] = generator.normalize_address_val(addr)

            if 'Type' in new_row:
                new_row['Type'] = self.normalize_type(new_row['Type'])

            if 'Name' not in new_row or not new_row['Name']:
                continue

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
