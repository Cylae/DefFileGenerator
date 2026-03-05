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
        if not action:
            return '1'
        a = str(action).upper().strip()
        if a in ['R', 'READ', '4']:
            return '4'
        if a in ['RW', 'W', 'WRITE', '1']:
            return '1'
        return a

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        data = []
        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                content = f.read(1024)
                f.seek(0)
                delimiter = ','
                for d in [',', ';', '\t']:
                    if d in content:
                        delimiter = d
                        break
                reader = csv.DictReader(f, delimiter=delimiter)
                data = [row for row in reader]
        except Exception as e:
            logging.error(f"Error reading CSV: {e}")
        return [data] if data else []

    def extract_from_xml(self, filepath):
        logging.info(f"Extracting from XML: {filepath}")
        try:
            import pandas as pd
            df = pd.read_xml(filepath)
            return [df.to_dict(orient='records')]
        except ImportError:
            try:
                import lxml.etree as ET
                tree = ET.parse(filepath)
                root = tree.getroot()
                data = []
                for child in root:
                    data.append({c.tag: c.text for c in child})
                return [data]
            except Exception as e:
                logging.error(f"Error reading XML: {e}")
                return []
        except Exception as e:
            logging.error(f"Error reading XML with pandas: {e}")
            return []

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
        logging.info(f"Extracting from Excel: {filepath}")
        wb = openpyxl.load_workbook(filepath, data_only=True)

        sheets_to_process = []
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                logging.error(f"Sheet '{sheet_name}' not found in {filepath}")
                return []
            sheets_to_process.append(wb[sheet_name])
        else:
            sheets_to_process = wb.worksheets

        all_tables = []
        for ws in sheets_to_process:
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
            if table_data:
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
                if isinstance(pages, int):
                    target_pages = [pdf.pages[pages-1]]
                elif isinstance(pages, list):
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
                    if table_data:
                        all_tables.append(table_data)
        return all_tables

    def map_and_clean(self, raw_data):
        if raw_data and isinstance(raw_data[0], dict):
            raw_data = [raw_data]

        all_mapped_data = []
        generator = Generator()

        for table in raw_data:
            if not table:
                continue

            first_row = table[0]
            standard_cols_mapping = {}
            used_src_cols = set()

            for target, source in self.mapping.items():
                if source in first_row:
                    standard_cols_mapping[target] = source
                    used_src_cols.add(source)

            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag', 'Factor', 'ScaleFactor']
            for target in detection_order:
                if target in standard_cols_mapping:
                    continue
                for src_col in first_row.keys():
                    if src_col in used_src_cols:
                        continue
                    src_lower = str(src_col).lower()
                    if any(p in src_lower for p in self.COLUMN_MAPPING.get(target, [])):
                        standard_cols_mapping[target] = src_col
                        used_src_cols.add(src_col)
                        break

            for row in table:
                new_row = {}
                for target, source in standard_cols_mapping.items():
                    val = row.get(source)
                    if val is not None:
                        new_row[target] = str(val).strip()

                name = new_row.get('Name')
                addr = new_row.get('Address')
                if not name and not addr:
                    continue

                if addr:
                    addr_str = str(addr).strip()
                    if ',' in addr_str and '.' not in addr_str:
                        addr_str = addr_str.replace(',', '')
                    match = re.search(r'(0x[0-9A-Fa-f]+|-?\d+(_\d+)*)', addr_str)
                    if match:
                        new_row['Address'] = match.group(1)

                if 'Factor' in new_row:
                    factor = str(new_row['Factor'])
                    if '/' in factor:
                        try:
                            p = factor.split('/')
                            new_row['Factor'] = str(float(p[0]) / float(p[1]))
                        except (ValueError, ZeroDivisionError):
                            new_row['Factor'] = '1'

                if 'Action' in new_row:
                    new_row['Action'] = self.normalize_action(new_row['Action'])

                if 'RegisterType' not in new_row:
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

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
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
