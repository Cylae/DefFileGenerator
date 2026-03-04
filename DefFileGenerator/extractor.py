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

    def __init__(self, mapping=None):
        self.mapping = mapping or {}

    def normalize_action(self, action):
        """Normalizes action strings to Webdyn codes."""
        if action is None or action == '':
            return '1'
        a = str(action).upper().strip()
        if a in ['R', 'READ']:
            return '4'
        if a in ['RW', 'W', 'WRITE']:
            return '1'
        return a

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
        logging.info(f"Extracting from Excel: {filepath}")
        wb = openpyxl.load_workbook(filepath, data_only=True)
        sheets = [sheet_name] if sheet_name else wb.sheetnames

        all_tables = []
        for s_name in sheets:
            if s_name not in wb.sheetnames:
                continue
            ws = wb[s_name]
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
        delimiters = [',', ';', '\t']
        for delim in delimiters:
            try:
                with open(filepath, 'r', encoding='utf-8-sig') as f:
                    content = f.read(4096)
                    if delim not in content:
                        continue
                    f.seek(0)
                    reader = csv.DictReader(f, delimiter=delim)
                    data = list(reader)
                    if data and len(reader.fieldnames) > 1:
                        return [data]
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
            logging.debug(f"Pandas read_xml failed, trying etree: {e}")
            try:
                df = pd.read_xml(filepath, parser='etree')
                return [df.to_dict(orient='records')]
            except Exception as e2:
                logging.error(f"Error extracting from XML: {e2}")
        return []

    def map_and_clean(self, tables):
        """Processes multiple tables and normalizes them."""
        # Handle backward compatibility: if a single table (list of dicts) is passed
        if tables and isinstance(tables[0], dict):
            tables = [tables]

        generator = Generator()
        all_mapped_data = []

        for table in tables:
            if not table:
                continue

            first_row = table[0]
            standard_cols_mapping = {}
            assigned_src_cols = set()

            # Explicitly mapped columns
            for target, source in self.mapping.items():
                if source in first_row:
                    standard_cols_mapping[target] = source
                    assigned_src_cols.add(source)

            # Heuristic matching
            # Priority: RegisterType > Address > Name > Type > Unit > Action > Tag
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag', 'Factor', 'ScaleFactor']
            for target in detection_order:
                if target in standard_cols_mapping:
                    continue
                for k in first_row.keys():
                    if k in assigned_src_cols:
                        continue
                    k_lower = str(k).lower()
                    for pattern in self.COLUMN_MAPPING.get(target, []):
                        if pattern in k_lower:
                            standard_cols_mapping[target] = k
                            assigned_src_cols.add(k)
                            break
                    if target in standard_cols_mapping:
                        break

            for row in table:
                new_row = {}
                for target, source in standard_cols_mapping.items():
                    new_row[target] = row[source]

                # Copy unassigned columns
                for k, v in row.items():
                    if k not in assigned_src_cols and k not in new_row:
                        new_row[k] = v

                # Address Cleaning
                if 'Address' in new_row and new_row['Address']:
                    addr = str(new_row['Address']).strip()
                    if '_' in addr:
                        parts = addr.split('_')
                        norm_parts = [generator.normalize_address_val(p) for p in parts]
                        new_row['Address'] = '_'.join(norm_parts)
                    else:
                        new_row['Address'] = generator.normalize_address_val(addr)

                # Skip rows without Name or Address
                if ('Name' not in new_row or not new_row['Name']) and ('Address' not in new_row or not new_row['Address']):
                    continue

                if 'Name' not in new_row or not new_row['Name']:
                    new_row['Name'] = f"Register {new_row.get('Address', 'Unknown')}"

                if 'RegisterType' not in new_row:
                    new_row['RegisterType'] = 'Holding Register'

                # Data type and Action normalization will be handled by Generator.process_rows
                # but we can do a basic pass for Action if it's there
                if 'Action' in new_row:
                    new_row['Action'] = self.normalize_action(new_row['Action'])

                all_mapped_data.append(new_row)

        return all_mapped_data

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(description='Extract registers from documentation.')
    parser.add_argument('input_file', help='Path to PDF, Excel, CSV, or XML.')
    parser.add_argument('-o', '--output', help='Output CSV.')
    parser.add_argument('--mapping', help='JSON mapping file.')
    parser.add_argument('--sheet', help='Excel sheet name.')
    parser.add_argument('--pages', help='PDF pages (comma separated).')

    args = parser.parse_args()
    mapping = {}
    if args.mapping:
        with open(args.mapping, 'r') as f:
            mapping = json.load(f)

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
