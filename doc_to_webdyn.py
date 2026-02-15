#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
import csv
import math

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

from DefFileGenerator.def_gen import Generator

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
    'float': 'F32',
    'f32': 'F32',
    'float32': 'F32',
    'double': 'F64',
    'f64': 'F64',
    'float64': 'F64',
}

def is_na(val):
    if HAS_PANDAS:
        return pd.isna(val)
    return val is None or val == '' or (isinstance(val, float) and math.isnan(val))

class MockDF:
    def __init__(self, rows, columns):
        self.rows = rows
        self.columns = columns
    def iterrows(self):
        for i, row in enumerate(self.rows):
            yield i, row

def find_column(df_columns, target_key):
    for col in df_columns:
        col_lower = str(col).lower()
        for pattern in COLUMN_MAPPING[target_key]:
            if pattern in col_lower:
                return col
    return None

def normalize_address(addr):
    if is_na(addr):
        return ''
    addr_str = str(addr).strip()
    # Handle formats like 40,001
    if ',' in addr_str and '.' not in addr_str:
        addr_str = addr_str.replace(',', '')

    # Support Address_Length and Address_Start_Bit formats (e.g. 30001_10 or 30001_0_1)
    # or simple dec/hex addresses.
    # Try to extract the first number or hex string found
    match = re.search(r'(0x[0-9A-Fa-f]+|\d+(_\d+)*)', addr_str)
    if match:
        return match.group(1)

    return addr_str

def normalize_action(action):
    if is_na(action):
        return '1'
    a = str(action).upper().strip()
    if a == 'R':
        return '4'
    if a == 'RW' or a == 'W':
        return '1'
    return a

def normalize_type(dtype):
    if is_na(dtype):
        return 'U16'
    dtype_str = str(dtype).lower().strip()
    for key, val in TYPE_MAPPING.items():
        if key in dtype_str:
            return val
    # Clean up common characters like () or space
    dtype_str = re.sub(r'[^a-z0-9_]+', '', dtype_str)
    return dtype_str.upper() if dtype_str else 'U16'

def main():
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', help='Path to the manufacturer documentation file (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer', required=True, help='Manufacturer name')
    parser.add_argument('--model', required=True, help='Model name')
    parser.add_argument('-o', '--output', help='Output filename (default: auto-generated)')
    parser.add_argument('--protocol', default='modbusRTU', help='Protocol name (default: modbusRTU)')
    parser.add_argument('--category', default='Inverter', help='Device category (default: Inverter)')
    parser.add_argument('--sheet', help='Excel sheet name (processes all if not specified)')
    parser.add_argument('--address-offset', type=int, default=0, help='Offset to subtract from input addresses.')
    parser.add_argument('-v', '--verbose', action='store_true', help='Show detailed processing information')

    args = parser.parse_args()

    # Configure logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=log_level, format='%(levelname)s: %(message)s')

    logging.info(f"Processing {args.input_file} for {args.manufacturer} {args.model}")

    ext = os.path.splitext(args.input_file)[1].lower()
    dataframes = []

    if ext == '.csv':
        dataframes = load_csv(args.input_file)
    elif ext in ['.xlsx', '.xls']:
        dataframes = load_excel(args.input_file, args.sheet)
    elif ext == '.xml':
        dataframes = load_xml(args.input_file)
    elif ext == '.pdf':
        dataframes = load_pdf(args.input_file)
    else:
        logging.error(f"Unsupported file extension: {ext}")
        sys.exit(1)

    if not dataframes:
        logging.error("No data could be extracted from the file.")
        sys.exit(1)

    logging.info(f"Extracted {len(dataframes)} table(s) from the file.")

    all_extracted_rows = []

    for df in dataframes:
        # Clean column names
        df.columns = [str(c).strip() for c in df.columns]

        # Identify columns
        col_map = {}
        assigned_cols = set()

        # Priority order for detection to avoid misidentification (e.g. RegisterType as Type)
        detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Scale', 'Action']
        for key in detection_order:
            if key not in COLUMN_MAPPING:
                continue
            for col in df.columns:
                if col in assigned_cols:
                    continue
                col_lower = str(col).lower()
                for pattern in COLUMN_MAPPING[key]:
                    if pattern in col_lower:
                        col_map[key] = col
                        assigned_cols.add(col)
                        break
                if key in col_map:
                    break

        if 'Address' not in col_map and 'Name' not in col_map:
             logging.debug("Skipping table as neither Address nor Name columns found.")
             continue

        for _, row in df.iterrows():
            addr_raw = row.get(col_map.get('Address')) if 'Address' in col_map else None
            name_raw = row.get(col_map.get('Name')) if 'Name' in col_map else None

            if is_na(addr_raw) and is_na(name_raw):
                continue

            # Skip header-like rows
            if str(addr_raw).lower() in COLUMN_MAPPING['Address'] or str(name_raw).lower() in COLUMN_MAPPING['Name']:
                continue

            addr = normalize_address(addr_raw)
            name = str(name_raw) if not is_na(name_raw) else f"Register {addr}"

            dtype_raw = row.get(col_map.get('Type')) if 'Type' in col_map else 'U16'
            unit_raw = row.get(col_map.get('Unit')) if 'Unit' in col_map else ''
            scale_raw = row.get(col_map.get('Scale')) if 'Scale' in col_map else '1'
            reg_type_raw = row.get(col_map.get('RegisterType')) if 'RegisterType' in col_map else 'Holding Register'
            action_raw = row.get(col_map.get('Action')) if 'Action' in col_map else '1'

            action = normalize_action(action_raw)

            # Clean scale (sometimes it's "1/10" or "0.1")
            scale = str(scale_raw)
            if '/' in scale:
                try:
                    parts = scale.split('/')
                    scale = str(float(parts[0]) / float(parts[1]))
                except (ValueError, ZeroDivisionError):
                    scale = '1'

            extracted_row = {
                'Name': name,
                'Tag': '', # Will be automatically generated by Generator
                'RegisterType': str(reg_type_raw) if not is_na(reg_type_raw) else 'Holding Register',
                'Address': addr,
                'Type': normalize_type(dtype_raw),
                'Factor': scale if scale and scale != 'nan' else '1',
                'Offset': '0',
                'Unit': str(unit_raw) if not is_na(unit_raw) and str(unit_raw) != 'nan' else '',
                'Action': action
            }
            all_extracted_rows.append(extracted_row)

    if not all_extracted_rows:
        logging.error("No registers could be extracted from the tables.")
        sys.exit(1)

    logging.info(f"Successfully extracted {len(all_extracted_rows)} registers.")

    # Use DefFileGenerator logic to process and validate
    generator = Generator(address_offset=args.address_offset)
    processed_rows = generator.process_rows(all_extracted_rows)

    # Determine output filename
    output_file = args.output
    if not output_file:
        safe_mfg = re.sub(r'[^a-zA-Z0-9]', '_', args.manufacturer).lower()
        safe_model = re.sub(r'[^a-zA-Z0-9]', '_', args.model).lower()
        output_file = f"{safe_mfg}_{safe_model}_definition.csv"

    # Write Output in WebdynSunPM format
    try:
        with open(output_file, 'w', newline='', encoding='utf-8') as outfile:
            # Prepare output header row
            header_row = [
                args.protocol,
                args.category,
                args.manufacturer,
                args.model,
                '', # ForcedWriteCode
                '', '', '', '', '', ''
            ]

            writer = csv.writer(outfile, delimiter=';', lineterminator='\n')
            writer.writerow(header_row)

            for index, row in enumerate(processed_rows, start=1):
                data_row = [
                    str(index),
                    row['Info1'],
                    row['Info2'],
                    row['Info3'],
                    row['Info4'],
                    row['Name'],
                    row['Tag'],
                    row['CoefA'],
                    row['CoefB'],
                    row['Unit'],
                    row['Action']
                ]
                writer.writerow(data_row)

        logging.info(f"Definition file successfully generated at {output_file}")
    except Exception as e:
        logging.error(f"Error writing output file: {e}")
        sys.exit(1)

def load_csv(filepath):
    if HAS_PANDAS:
        for delimiter in [',', ';', '\t']:
            try:
                df = pd.read_csv(filepath, sep=delimiter, encoding='utf-8-sig')
                if len(df.columns) > 1:
                    return [df]
            except Exception:
                continue
    else:
        # Fallback to standard csv module
        for delimiter in [',', ';', '\t']:
             try:
                 with open(filepath, 'r', encoding='utf-8-sig') as f:
                     reader = csv.DictReader(f, delimiter=delimiter)
                     rows = list(reader)
                     if rows and len(reader.fieldnames) > 1:
                         return [MockDF(rows, reader.fieldnames)]
             except Exception:
                 continue
    return []

def load_excel(filepath, sheet_name=None):
    if not HAS_PANDAS:
        logging.error("pandas and openpyxl are required for Excel processing. Please install them.")
        return []
    try:
        if sheet_name:
            df = pd.read_excel(filepath, sheet_name=sheet_name)
            return [df]
        else:
            excel_file = pd.ExcelFile(filepath)
            return [excel_file.parse(sheet) for sheet in excel_file.sheet_names]
    except Exception as e:
        logging.error(f"Error loading Excel file: {e}")
        return []

def load_xml(filepath):
    if not HAS_PANDAS:
        logging.error("pandas and lxml are required for XML processing. Please install them.")
        return []
    try:
        # Try with default parser (usually lxml if installed)
        df = pd.read_xml(filepath)
        return [df]
    except Exception as e:
        logging.debug(f"Default XML parser failed, trying etree: {e}")
        try:
            # Fallback to etree which is in the standard library
            df = pd.read_xml(filepath, parser='etree')
            return [df]
        except Exception as e2:
            logging.error(f"Error loading XML file: {e2}")
            return []

def load_pdf(filepath):
    if not HAS_PDFPLUMBER:
        logging.error("pdfplumber is required for PDF processing. Please install it.")
        return []
    dfs = []
    try:
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and len(table) > 1:
                        # Use first row as header if it looks like one
                        headers = table[0]
                        # Clean headers (remove newlines)
                        headers = [str(h).replace('\n', ' ') if h else f"Col{i}" for i, h in enumerate(headers)]
                        if HAS_PANDAS:
                            df = pd.DataFrame(table[1:], columns=headers)
                        else:
                            # Mock DF for PDF as well if pandas missing but pdfplumber present
                            rows = []
                            for row_data in table[1:]:
                                rows.append(dict(zip(headers, row_data)))
                            df = MockDF(rows, headers)
                        dfs.append(df)
    except Exception as e:
        logging.error(f"Error loading PDF file: {e}")
    return dfs

if __name__ == "__main__":
    main()
