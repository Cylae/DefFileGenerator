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

        # Apply known mappings
        for key, val in self.TYPE_MAPPING.items():
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

        return t.upper()

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
        logging.info(f"Extracting from Excel: {filepath}")
        wb = openpyxl.load_workbook(filepath, data_only=True)

        results = []
        if sheet_name:
            sheets = [sheet_name] if sheet_name in wb.sheetnames else []
        else:
            sheets = wb.sheetnames

        for name in sheets:
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
                results.append(data)
        return results

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        logging.info(f"Extracting from PDF: {filepath}")
        results = []
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
                    data = []
                    for row in table[1:]:
                        row_data = {}
                        for i, cell in enumerate(row):
                            if i < len(headers):
                                val = str(cell).replace('\n', ' ').strip() if cell else ""
                                row_data[headers[i]] = val
                        data.append(row_data)
                    if data:
                        results.append(data)
        return results

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        for delimiter in [',', ';', '\t']:
            try:
                with open(filepath, 'r', encoding='utf-8-sig') as f:
                    sample = f.read(2048)
                    f.seek(0)
                    if delimiter not in sample:
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
             logging.error("pandas is required for XML extraction.")
             return []
        logging.info(f"Extracting from XML: {filepath}")
        try:
            df = pd.read_xml(filepath)
            return [df.to_dict(orient='records')]
        except Exception as e:
            logging.error(f"Error extracting from XML: {e}")
            return []

    def map_and_clean(self, tables):
        """Processes one or more tables of raw data."""
        if not tables:
            return []

        # If a single table was passed (list of dicts), wrap it in a list
        if isinstance(tables, list) and len(tables) > 0 and isinstance(tables[0], dict):
            tables = [tables]

        final_data = []
        generator = Generator()

        for raw_data in tables:
            if not raw_data:
                continue

            first_row = raw_data[0]
            src_cols = list(first_row.keys())
            standard_cols_mapping = {}
            used_src_cols = set()

            # 1. Explicit mapping from config
            for target, source in self.mapping.items():
                if source in src_cols:
                    standard_cols_mapping[target] = source
                    used_src_cols.add(source)

            # 2. Heuristic mapping based on COLUMN_MAPPING
            # Priority order for detection
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Factor', 'ScaleFactor', 'Tag', 'Action']
            for target in detection_order:
                if target in standard_cols_mapping:
                    continue

                # Try to find a match in src_cols
                for col in src_cols:
                    if col in used_src_cols:
                        continue

                    col_lower = str(col).lower()
                    patterns = self.COLUMN_MAPPING.get(target, [target.lower()])

                    if any(p in col_lower for p in patterns):
                        standard_cols_mapping[target] = col
                        used_src_cols.add(col)
                        break

            for row in raw_data:
                new_row = {}
                for target, source in standard_cols_mapping.items():
                    new_row[target] = row[source]

                # Fill in unmapped columns
                for k, v in row.items():
                    if k not in used_src_cols:
                        new_row[k] = v

                # Name is mandatory
                if 'Name' not in new_row or not str(new_row.get('Name', '')).strip():
                    continue

                # Clean Address
                addr_raw = new_row.get('Address')
                if addr_raw:
                    addr_str = str(addr_raw).strip()
                    # Delegate final normalization to Generator
                    if '_' in addr_str:
                        parts = addr_str.split('_')
                        norm_parts = [generator.normalize_address_val(p) for p in parts]
                        new_row['Address'] = '_'.join(norm_parts)
                    else:
                        new_row['Address'] = generator.normalize_address_val(addr_str)
                else:
                    # If no address, skip unless it's a header-like row (which we should skip anyway)
                    continue

                # Clean Type
                if 'Type' in new_row:
                    new_row['Type'] = self.normalize_type(new_row['Type'])
                else:
                    new_row['Type'] = 'U16'

                # Clean Factor (handle fractions)
                if 'Factor' in new_row:
                    f_val = str(new_row['Factor'])
                    if '/' in f_val:
                        try:
                            parts = f_val.split('/')
                            new_row['Factor'] = str(float(parts[0]) / float(parts[1]))
                        except (ValueError, ZeroDivisionError):
                            new_row['Factor'] = '1'

                final_data.append(new_row)

        return final_data

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

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
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
