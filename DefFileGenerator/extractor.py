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
        'uint16': 'U16', 'int16': 'I16',
        'uint32': 'U32', 'int32': 'I32',
        'uint64': 'U64', 'int64': 'I64',
        'float': 'F32', 'f32': 'F32', 'float32': 'F32',
        'double': 'F64', 'f64': 'F64', 'float64': 'F64',
        'string': 'STRING', 'bits': 'BITS'
    }

    def __init__(self, mapping=None):
        self.mapping = mapping or {}

    def normalize_type(self, t):
        # Delegate to Generator for consistency
        return Generator().normalize_type(t)

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
        logging.info(f"Extracting from Excel: {filepath}")
        wb = openpyxl.load_workbook(filepath, data_only=True)
        sheets_to_process = [sheet_name] if sheet_name else wb.sheetnames

        tables = []
        for name in sheets_to_process:
            if name not in wb.sheetnames:
                logging.error(f"Sheet '{name}' not found in {filepath}")
                continue
            ws = wb[name]
            rows = list(ws.rows)
            if not rows:
                continue

            headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]
            data = []
            for row in rows[1:]:
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
        if HAS_PANDAS:
            for delimiter in [',', ';', '\t']:
                try:
                    df = pd.read_csv(filepath, sep=delimiter, encoding='utf-8-sig')
                    if len(df.columns) > 1:
                        return [df.to_dict(orient='records')]
                except Exception:
                    continue
        else:
            for delimiter in [',', ';', '\t']:
                 try:
                     with open(filepath, 'r', encoding='utf-8-sig') as f:
                         reader = csv.DictReader(f, delimiter=delimiter)
                         rows = list(reader)
                         if rows and len(reader.fieldnames) > 1:
                             return [rows]
                 except Exception:
                     continue
        return []

    def extract_from_xml(self, filepath):
        if not HAS_PANDAS:
            logging.error("pandas and lxml are required for XML processing.")
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

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        logging.info(f"Extracting from PDF: {filepath}")
        tables = []
        with pdfplumber.open(filepath) as pdf:
            if pages is None:
                target_pages = pdf.pages
            else:
                target_pages = [pdf.pages[i-1] for i in pages if 0 < i <= len(pdf.pages)]

            for page in target_pages:
                page_tables = page.extract_tables()
                for table in page_tables:
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
                        tables.append(data)
        return tables

    def map_and_clean_table(self, table_data):
        if not table_data:
            return []

        # Find first non-empty row to determine headers
        first_row = {}
        for row in table_data:
            if any(v is not None and str(v).strip() != "" for v in row.values()):
                first_row = row
                break
        if not first_row: return []

        standard_cols_mapping = {}
        used_src_cols = set()

        # 1. Explicitly mapped columns from config
        for target, source in self.mapping.items():
            if source in first_row:
                standard_cols_mapping[target] = source
                used_src_cols.add(source)

        # 2. Heuristic match for standard columns
        detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Factor', 'ScaleFactor', 'Tag', 'Action']
        for target in detection_order:
            if target in standard_cols_mapping:
                continue
            patterns = self.COLUMN_MAPPING.get(target, [])
            for src_col in first_row.keys():
                if src_col in used_src_cols:
                    continue
                src_col_lower = str(src_col).lower()
                for pattern in patterns:
                    if pattern in src_col_lower:
                        standard_cols_mapping[target] = src_col
                        used_src_cols.add(src_col)
                        break
                if target in standard_cols_mapping:
                    break

        generator = Generator()
        mapped_rows = []
        for row in table_data:
            if not any(v is not None and str(v).strip() != "" for v in row.values()):
                continue
            new_row = {}
            for target, source in standard_cols_mapping.items():
                val = row.get(source)
                if val is not None:
                    new_row[target] = val
            for k, v in row.items():
                if k not in used_src_cols and k not in new_row:
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
                new_row['Type'] = generator.normalize_type(new_row['Type'])
            if 'Action' in new_row:
                new_row['Action'] = generator.normalize_action(new_row['Action'])
            if 'Name' not in new_row or not new_row['Name']:
                if 'Address' in new_row and new_row['Address']:
                    new_row['Name'] = f"Register {new_row['Address']}"
                else: continue
            if 'RegisterType' not in new_row:
                new_row['RegisterType'] = 'Holding Register'
            mapped_rows.append(new_row)
        return mapped_rows

    def map_and_clean(self, tables):
        if not tables:
            return []

        # Detect if we were passed a single table (list of dicts) or list of tables
        # If it's a list of dicts, it's a single table.
        # If it's a list of lists, it's multiple tables.
        if isinstance(tables, list) and len(tables) > 0 and isinstance(tables[0], dict):
            return self.map_and_clean_table(tables)

        all_mapped = []
        for table in tables:
            if isinstance(table, list):
                all_mapped.extend(self.map_and_clean_table(table))
        return all_mapped

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
        tables = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = [int(p.strip()) for p in args.pages.split(',')] if args.pages else None
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
