#!/usr/bin/env python3
import argparse
import csv
import json
import logging
import os
import re
import sys
import math

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
        'Action': ['action', 'access']
    }

    TYPE_MAPPING = {
        'uint16': 'U16',
        'int16': 'I16',
        'uint32': 'U32',
        'int32': 'I32',
        'uint64': 'U64',
        'int64': 'I64',
        'u16': 'U16',
        'i16': 'I16',
        'u32': 'U32',
        'i32': 'I32',
        'f32': 'F32',
        'float32': 'F32',
        'float': 'F32',
        'double': 'F64',
        'f64': 'F64',
        'float64': 'F64',
        'string': 'STRING',
        'bits': 'BITS'
    }

    def __init__(self, mapping=None):
        self.mapping = mapping or {}

    def normalize_type(self, t):
        if t is None or (isinstance(t, float) and math.isnan(t)):
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

        # Clean up common characters like () or space
        t_str = re.sub(r'[^a-z0-9_]+', '', t_str)
        return t_str.upper() if t_str else 'U16'

    def normalize_action(self, action):
        if action is None or (isinstance(action, float) and math.isnan(action)):
            return '1'
        act_str = str(action).strip().upper()
        if act_str in ['R', 'READ', '4']:
            return '4'
        if act_str in ['RW', 'W', 'WRITE', '1']:
            return '1'
        # If it's already a valid numeric code, keep it
        if act_str in ['0', '1', '2', '3', '4', '6', '7', '8', '9']:
            return act_str
        return '1' # Default

    def extract_from_excel(self, filepath, sheet_name=None):
        if HAS_PANDAS:
            logging.info(f"Extracting from Excel (via pandas): {filepath}")
            try:
                if sheet_name:
                    df = pd.read_excel(filepath, sheet_name=sheet_name)
                    return [df.to_dict('records')]
                else:
                    excel_file = pd.ExcelFile(filepath)
                    return [excel_file.parse(sheet).to_dict('records') for sheet in excel_file.sheet_names]
            except Exception as e:
                logging.error(f"Error loading Excel file via pandas: {e}")

        if not HAS_OPENPYXL:
            logging.error("pandas or openpyxl is required for Excel extraction.")
            return []

        logging.info(f"Extracting from Excel (via openpyxl): {filepath}")
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            if sheet_name:
                if sheet_name not in wb.sheetnames:
                    logging.error(f"Sheet '{sheet_name}' not found in {filepath}")
                    return []
                sheets = [wb[sheet_name]]
            else:
                sheets = wb.worksheets

            all_data = []
            for ws in sheets:
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
                all_data.append(data)
            return all_data
        except Exception as e:
            logging.error(f"Error loading Excel file via openpyxl: {e}")
            return []

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        logging.info(f"Extracting from PDF: {filepath}")
        all_tables = []
        try:
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
                        all_tables.append(data)
        except Exception as e:
            logging.error(f"Error loading PDF file: {e}")
        return all_tables

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        for delimiter in [',', ';', '\t']:
            try:
                if HAS_PANDAS:
                    df = pd.read_csv(filepath, sep=delimiter, encoding='utf-8-sig')
                    if len(df.columns) > 1:
                        return [df.to_dict('records')]
                else:
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
            logging.error("pandas is required for XML extraction.")
            return []
        logging.info(f"Extracting from XML: {filepath}")
        try:
            df = pd.read_xml(filepath)
            return [df.to_dict('records')]
        except Exception as e:
            logging.debug(f"Default XML parser failed, trying etree: {e}")
            try:
                df = pd.read_xml(filepath, parser='etree')
                return [df.to_dict('records')]
            except Exception as e2:
                logging.error(f"Error loading XML file: {e2}")
                return []

    def map_and_clean(self, raw_data_list):
        if not raw_data_list:
            return []

        # Handle single table or list of tables
        if isinstance(raw_data_list, list) and len(raw_data_list) > 0 and isinstance(raw_data_list[0], dict):
             raw_data_list = [raw_data_list]

        all_mapped_data = []
        generator = Generator()

        for table in raw_data_list:
            if not table:
                continue

            first_row = table[0]
            source_cols = list(first_row.keys())
            standard_cols_mapping = {}
            used_src_cols = set()

            # Explicitly mapped columns from config
            for target, source in self.mapping.items():
                if source in first_row:
                    standard_cols_mapping[target] = source
                    used_src_cols.add(source)

            # Heuristic mapping for standard columns
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Scale', 'Action', 'Tag']
            for target in detection_order:
                if target in standard_cols_mapping:
                    continue
                patterns = self.COLUMN_MAPPING.get(target, [target.lower()])
                for src_col in source_cols:
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

            for row in table:
                new_row = {}
                for target, source in standard_cols_mapping.items():
                    if source in row:
                        val = row[source]
                        if target == 'Scale':
                            scale = str(val) if val is not None and not (isinstance(val, float) and math.isnan(val)) else '1'
                            if '/' in scale:
                                try:
                                    parts = scale.split('/')
                                    scale = str(float(parts[0]) / float(parts[1]))
                                except (ValueError, ZeroDivisionError):
                                    scale = '1'
                            new_row['Factor'] = scale
                        else:
                            new_row[target] = val

                for k, v in row.items():
                    if k not in used_src_cols and k not in new_row:
                        new_row[k] = v

                # Clean Address
                if 'Address' in new_row and new_row['Address'] is not None:
                    addr = str(new_row['Address']).strip()
                    if addr and addr.lower() != 'nan':
                        if '_' in addr:
                            parts = addr.split('_')
                            norm_parts = [generator.normalize_address_val(p) for p in parts]
                            new_row['Address'] = '_'.join(norm_parts)
                        else:
                            new_row['Address'] = generator.normalize_address_val(addr)
                    else:
                        new_row['Address'] = ''

                # Clean Type
                if 'Type' in new_row:
                    new_row['Type'] = self.normalize_type(new_row['Type'])

                # Clean Action
                if 'Action' in new_row:
                    new_row['Action'] = self.normalize_action(new_row['Action'])

                # Ensure mandatory fields
                if 'Name' not in new_row or not new_row['Name'] or str(new_row['Name']).lower() == 'nan':
                    continue

                if 'RegisterType' not in new_row:
                    new_row['RegisterType'] = 'Holding Register'

                all_mapped_data.append(new_row)

        return all_mapped_data

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(description='Extract register information from documentation files.')
    parser.add_argument('input_file', help='Path to the source file.')
    parser.add_argument('-o', '--output', help='Path to the output CSV file.')
    parser.add_argument('--mapping', help='JSON file containing column mapping.')
    parser.add_argument('--sheet', help='Excel sheet name to extract from.')
    parser.add_argument('--pages', help='PDF pages to extract from (comma separated).')
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

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xls']:
        raw_data = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = [int(p.strip()) for p in args.pages.split(',')] if args.pages else None
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
