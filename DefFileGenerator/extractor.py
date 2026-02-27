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
        'Action': ['action', 'access'],
        'Tag': ['tag']
    }

    def __init__(self, mapping=None):
        self.mapping = mapping or {}
        self.generator = Generator()

    def normalize_type(self, t):
        return self.generator.normalize_type(t)

    def normalize_action(self, a):
        return self.generator.normalize_action(a)

    def extract_from_csv(self, filepath):
        """Extracts from CSV with automatic delimiter detection."""
        logging.info(f"Extracting from CSV: {filepath}")
        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                content = f.read(2048)
                f.seek(0)
                sniffer = csv.Sniffer()
                dialect = sniffer.sniff(content, delimiters=',;\t')

                reader = csv.DictReader(f, dialect=dialect)
                rows = list(reader)
                return [rows]
        except Exception as e:
            logging.error(f"Error extracting from CSV: {e}")
            return []

    def extract_from_xml(self, filepath):
        """Extracts from XML using pandas."""
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
                logging.error(f"Error extracting from XML: {e2}")
                return []

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
        logging.info(f"Extracting from Excel: {filepath}")
        wb = openpyxl.load_workbook(filepath, data_only=True)
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                logging.error(f"Sheet '{sheet_name}' not found in {filepath}")
                return []
            ws = wb[sheet_name]
        else:
            ws = wb.active

        all_sheets_data = []
        sheets_to_process = [sheet_name] if sheet_name else wb.sheetnames

        for name in sheets_to_process:
            if name not in wb.sheetnames:
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
            all_sheets_data.append(data)
        return all_sheets_data

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        logging.info(f"Extracting from PDF: {filepath}")
        all_tables_data = []
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
                    all_tables_data.append(data)
        return all_tables_data

    def map_and_clean(self, raw_data):
        if not raw_data:
            return []

        # Handle both single table (backward compatibility) and list of tables
        if isinstance(raw_data, list) and len(raw_data) > 0 and isinstance(raw_data[0], list):
            all_mapped_data = []
            for table in raw_data:
                all_mapped_data.extend(self._map_and_clean_single_table(table))
            return all_mapped_data
        else:
            return self._map_and_clean_single_table(raw_data)

    def _map_and_clean_single_table(self, raw_table):
        mapped_data = []
        if not raw_table:
            return []

        first_row = raw_table[0]
        standard_cols_mapping = {}
        assigned_keys = set()

        # Explicitly mapped columns
        for target, source in self.mapping.items():
            if source in first_row:
                standard_cols_mapping[target] = source
                assigned_keys.add(source)

        # Fuzzy match with priority
        detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Scale', 'Action', 'Tag']
        for target in detection_order:
            if target in standard_cols_mapping:
                continue
            patterns = self.COLUMN_MAPPING.get(target, [])
            for k in first_row.keys():
                if k in assigned_keys:
                    continue
                k_lower = str(k).lower()
                if any(p in k_lower for p in patterns):
                    standard_cols_mapping[target] = k
                    assigned_keys.add(k)
                    break

        for row in raw_table:
            new_row = {}
            for target, source in standard_cols_mapping.items():
                new_row[target] = row.get(source)

            # Extra columns
            for k, v in row.items():
                if k not in assigned_keys and k not in new_row:
                    new_row[k] = v

            if not new_row.get('Name') and not new_row.get('Address'):
                continue

            # Basic cleaning
            if new_row.get('Address'):
                addr = str(new_row['Address']).strip()
                # If it's a messy string, let Generator.normalize_address_val handle it later
                # but we can do a preliminary split for Addr_Len/Addr_Start_NbBits
                if '_' in addr:
                    parts = addr.split('_')
                    norm_parts = [self.generator.normalize_address_val(p) for p in parts]
                    new_row['Address'] = '_'.join(norm_parts)
                else:
                    new_row['Address'] = self.generator.normalize_address_val(addr)

            if new_row.get('Type'):
                new_row['Type'] = self.normalize_type(new_row['Type'])

            if new_row.get('Action'):
                new_row['Action'] = self.normalize_action(new_row['Action'])

            if not new_row.get('RegisterType'):
                new_row['RegisterType'] = 'Holding Register'

            # Rename Scale to Factor/ScaleFactor if needed
            if 'Scale' in new_row:
                scale_val = str(new_row['Scale'])
                if '/' in scale_val:
                    try:
                        p = scale_val.split('/')
                        new_row['Factor'] = str(float(p[0])/float(p[1]))
                    except:
                        new_row['Factor'] = '1'
                else:
                    new_row['Factor'] = scale_val

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
