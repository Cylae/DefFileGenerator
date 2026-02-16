#!/usr/bin/env python3
import argparse
import csv
import json
import logging
import os
import re
import sys
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

try:
    import lxml
    HAS_LXML = True
except ImportError:
    HAS_LXML = False

from DefFileGenerator.def_gen import Generator

class Extractor:
    COLUMN_MAPPING = {
        'RegisterType': ['register type', 'reg type', 'modbus type', 'registertype'],
        'Address': ['address', 'addr', 'offset', 'register', 'reg'],
        'Name': ['name', 'description', 'parameter', 'variable', 'signal'],
        'Type': ['data type', 'datatype', 'type', 'format'],
        'Unit': ['unit', 'units'],
        'Factor': ['scale', 'factor', 'multiplier', 'ratio'],
        'Action': ['action', 'access'],
        'ScaleFactor': ['scalefactor'],
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
        logging.info(f"Extracting from Excel: {filepath}")
        if HAS_PANDAS:
            try:
                if sheet_name:
                    df = pd.read_excel(filepath, sheet_name=sheet_name)
                    return df.to_dict(orient='records')
                else:
                    excel_file = pd.ExcelFile(filepath)
                    all_data = []
                    for sheet in excel_file.sheet_names:
                        df = excel_file.parse(sheet)
                        all_data.extend(df.to_dict(orient='records'))
                    return all_data
            except Exception as e:
                logging.error(f"Error loading Excel with pandas: {e}")

        if not HAS_OPENPYXL:
            logging.error("pandas or openpyxl is required for Excel extraction.")
            return []

        wb = openpyxl.load_workbook(filepath, data_only=True)
        sheets = [sheet_name] if sheet_name else wb.sheetnames
        data = []
        for sname in sheets:
            if sname not in wb.sheetnames:
                continue
            ws = wb[sname]
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
        return data

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        if HAS_PANDAS:
            for delimiter in [',', ';', '\t']:
                try:
                    df = pd.read_csv(filepath, sep=delimiter, encoding='utf-8-sig')
                    if len(df.columns) > 1:
                        return df.to_dict(orient='records')
                except Exception:
                    continue

        # Fallback to standard csv module
        for delimiter in [',', ';', '\t']:
            try:
                with open(filepath, 'r', encoding='utf-8-sig') as f:
                    reader = csv.DictReader(f, delimiter=delimiter)
                    rows = list(reader)
                    if rows and len(reader.fieldnames) > 1:
                        return rows
            except Exception:
                continue
        return []

    def extract_from_xml(self, filepath):
        logging.info(f"Extracting from XML: {filepath}")
        if not HAS_PANDAS:
            logging.error("pandas is required for XML extraction.")
            return []
        try:
            df = pd.read_xml(filepath)
            return df.to_dict(orient='records')
        except Exception as e:
            logging.debug(f"Default XML parser failed, trying etree: {e}")
            try:
                df = pd.read_xml(filepath, parser='etree')
                return df.to_dict(orient='records')
            except Exception as e2:
                logging.error(f"Error loading XML: {e2}")
                return []

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
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
        if not raw_data:
            return []

        mapped_data = []
        # Identify standard columns once
        first_row = raw_data[0]
        standard_cols_mapping = {}
        assigned_keys = set()

        # Explicitly mapped columns from config
        for target, source in self.mapping.items():
            if source in first_row:
                standard_cols_mapping[target] = source
                assigned_keys.add(source)

        # Fuzzy match for standard columns
        # Priority order to avoid misidentification
        detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Factor', 'Action', 'ScaleFactor', 'Tag']
        for target in detection_order:
            if target in standard_cols_mapping:
                continue
            for k in first_row.keys():
                if k in assigned_keys:
                    continue
                k_lower = str(k).lower()
                patterns = self.COLUMN_MAPPING.get(target, [])
                if any(p in k_lower for p in patterns):
                    standard_cols_mapping[target] = k
                    assigned_keys.add(k)
                    break

        generator = Generator()
        for row in raw_data:
            new_row = {}
            for target, source in standard_cols_mapping.items():
                val = row.get(source)
                # Map Scale/Factor to Factor
                if target == 'Factor':
                    new_row['Factor'] = val
                else:
                    new_row[target] = val

            # Fill in other columns
            for k, v in row.items():
                if k not in assigned_keys and k not in new_row:
                    new_row[k] = v

            # Ensure mandatory fields for def_gen
            if 'Name' not in new_row or not new_row['Name'] or str(new_row['Name']).lower() == 'nan':
                continue

            # Clean Address using Generator's logic
            if 'Address' in new_row and new_row['Address']:
                addr = str(new_row['Address']).strip()
                if '_' in addr:
                    parts = addr.split('_')
                    norm_parts = [generator.normalize_address_val(p) for p in parts]
                    new_row['Address'] = '_'.join(norm_parts)
                else:
                    new_row['Address'] = generator.normalize_address_val(addr)
            else:
                continue # Skip rows without address

            # Clean Type
            if 'Type' in new_row:
                new_row['Type'] = self.normalize_type(new_row['Type'])
            else:
                new_row['Type'] = 'U16'

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
