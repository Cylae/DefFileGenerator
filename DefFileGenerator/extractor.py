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
        'Action': ['action', 'access'],
        'Offset': ['offset']
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

        target_sheets = []
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                logging.error(f"Sheet '{sheet_name}' not found in {filepath}")
                return []
            target_sheets = [wb[sheet_name]]
        else:
            target_sheets = wb.worksheets

        all_sheets_data = []
        for ws in target_sheets:
            sheet_data = []
            rows = list(ws.rows)
            if not rows or len(rows) < 2:
                continue

            headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]

            for row in rows[1:]:
                row_data = {}
                has_data = False
                for i, cell in enumerate(row):
                    if i < len(headers):
                        val = cell.value
                        if val is not None:
                            has_data = True
                        row_data[headers[i]] = val
                if has_data:
                    sheet_data.append(row_data)
            if sheet_data:
                all_sheets_data.append(sheet_data)
        return all_sheets_data

    def extract_from_xml(self, filepath):
        logging.info(f"Extracting from XML: {filepath}")
        try:
            import pandas as pd
            # Try with default parser (usually lxml if installed)
            df = pd.read_xml(filepath)
            return [df.to_dict(orient='records')]
        except ImportError:
            logging.error("pandas is required for XML extraction.")
        except Exception as e:
            logging.debug(f"Default XML parser failed, trying etree: {e}")
            try:
                import pandas as pd
                df = pd.read_xml(filepath, parser='etree')
                return [df.to_dict(orient='records')]
            except Exception as e2:
                logging.error(f"Error loading XML file: {e2}")
        return []

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        # Use utf-8-sig to handle potential BOM
        for delimiter in [',', ';', '\t']:
            try:
                with open(filepath, 'r', encoding='utf-8-sig') as f:
                    # Check if delimiter exists in first line
                    first_line = f.readline()
                    if delimiter not in first_line and delimiter != ',':
                        continue
                    f.seek(0)
                    reader = csv.DictReader(f, delimiter=delimiter)
                    rows = list(reader)
                    if rows and len(reader.fieldnames) > 1:
                        # Clean fieldnames
                        cleaned_rows = []
                        for row in rows:
                            cleaned_row = {str(k).strip(): v for k, v in row.items()}
                            cleaned_rows.append(cleaned_row)
                        return [cleaned_rows]
            except Exception:
                continue

        # Final attempt with standard csv if above failed
        try:
             with open(filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                rows = list(reader)
                if rows:
                    return [rows]
        except Exception as e:
            logging.error(f"Error loading CSV file: {e}")

        return []

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

                    table_data = []
                    # Clean headers: remove newlines
                    headers = [str(c).replace('\n', ' ').strip() if c else "" for c in table[0]]

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

    def normalize_action(self, action):
        if action is None or action == "":
            return '1'
        a = str(action).upper().strip()
        if a in ['R', 'READ', '4']:
            return '4'
        if a in ['RW', 'W', 'WRITE', '1']:
            return '1'
        return a

    def map_and_clean(self, raw_data):
        """
        Maps source columns to standard names and cleans the data.
        raw_data can be List[dict] or List[List[dict]].
        Returns a flattened List[dict] of cleaned register data.
        """
        if not raw_data:
            return []

        # Ensure raw_data is a list of tables
        if isinstance(raw_data, list) and len(raw_data) > 0 and isinstance(raw_data[0], dict):
            tables = [raw_data]
        else:
            tables = raw_data

        all_mapped_rows = []
        generator = Generator()

        for table in tables:
            if not table:
                continue

            # Identify columns for this table
            first_row = table[0]
            col_map = {}
            assigned_src_cols = set()

            # 1. Use explicit mapping if provided
            for target, source in self.mapping.items():
                if source in first_row:
                    col_map[target] = source
                    assigned_src_cols.add(source)

            # 2. Heuristic mapping using COLUMN_MAPPING
            # Priority order for detection to avoid misidentification
            detection_order = [
                'RegisterType', 'Address', 'Name', 'Type', 'Unit',
                'Factor', 'ScaleFactor', 'Tag', 'Action', 'Offset'
            ]
            for target in detection_order:
                if target in col_map:
                    continue

                # Check each source column for matches in COLUMN_MAPPING patterns
                for src_col in first_row.keys():
                    if src_col in assigned_src_cols:
                        continue

                    src_col_lower = str(src_col).lower()
                    patterns = self.COLUMN_MAPPING.get(target, [])

                    for pattern in patterns:
                        if pattern in src_col_lower:
                            col_map[target] = src_col
                            assigned_src_cols.add(src_col)
                            break
                    if target in col_map:
                        break

            # 3. Process rows in this table
            for row in table:
                new_row = {}
                for target, source in col_map.items():
                    new_row[target] = row.get(source)

                # If Name and Address are missing, skip
                if not new_row.get('Name') and not new_row.get('Address'):
                    continue

                # Normalization
                if new_row.get('Address'):
                    addr = str(new_row['Address']).strip()
                    if '_' in addr:
                        parts = addr.split('_')
                        norm_parts = [generator.normalize_address_val(p) for p in parts]
                        new_row['Address'] = '_'.join(norm_parts)
                    else:
                        new_row['Address'] = generator.normalize_address_val(addr)

                new_row['Type'] = self.normalize_type(new_row.get('Type'))
                new_row['Action'] = self.normalize_action(new_row.get('Action'))

                # Default RegisterType
                if not new_row.get('RegisterType'):
                    new_row['RegisterType'] = 'Holding Register'

                # Clean Factor/Scale (sometimes it's "1/10" or "0.1")
                if new_row.get('Factor'):
                    factor_str = str(new_row['Factor'])
                    if '/' in factor_str:
                        try:
                            parts = factor_str.split('/')
                            new_row['Factor'] = str(float(parts[0]) / float(parts[1]))
                        except (ValueError, ZeroDivisionError):
                            new_row['Factor'] = '1'

                all_mapped_rows.append(new_row)

        return all_mapped_rows

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
