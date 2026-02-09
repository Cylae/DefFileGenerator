#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import pandas as pd
import pdfplumber
import re
import csv
from DefFileGenerator.def_gen import Generator

COLUMN_MAPPING = {
    'RegisterType': ['register type', 'reg type', 'info1'],
    'Address': ['register', 'address', 'addr', 'offset', 'reg'],
    'Name': ['name', 'description', 'parameter', 'variable', 'signal'],
    'Type': ['type', 'data type', 'format', 'datatype'],
    'Unit': ['unit', 'units'],
    'Scale': ['scale', 'factor', 'multiplier', 'ratio']
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

def find_column(df_columns, target_key):
    for col in df_columns:
        col_lower = str(col).lower()
        for pattern in COLUMN_MAPPING[target_key]:
            if pattern in col_lower:
                return col
    return None

def normalize_address(addr):
    if addr is None or pd.isna(addr):
        return ''
    addr_str = str(addr).strip()
    if addr_str.lower().startswith('0x'):
        try:
            return str(int(addr_str, 16))
        except ValueError:
            pass

    # Handle composite address formats (e.g., 30001_10 or 30001_0_1)
    if '_' in addr_str:
        parts = addr_str.split('_')
        if all(p.isdigit() for p in parts):
            return addr_str

    # Handle formats like 40,001
    if ',' in addr_str and '.' not in addr_str:
        addr_str = addr_str.replace(',', '')

    # Try to extract the first number found (to handle things like "40001 (Holding)")
    match = re.search(r'(\d+)', addr_str)
    if match:
        return match.group(1)

    return addr_str

def normalize_type(dtype):
    if dtype is None or pd.isna(dtype):
        return 'U16'
    dtype_str = str(dtype).lower().strip()
    for key, val in TYPE_MAPPING.items():
        if key in dtype_str:
            return val
    # Clean up common characters like () or space
    dtype_str = re.sub(r'[^a-z0-9_]+', '', dtype_str)
    return dtype_str.upper() if dtype_str else 'U16'

def generate_tag(name):
    if name is None or pd.isna(name):
        return ""
    tag = str(name).lower()
    # Replace non-alphanumeric with underscore
    tag = re.sub(r'[^a-z0-9]+', '_', tag)
    # Remove leading/trailing underscores
    return tag.strip('_')

def main():
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', help='Path to the manufacturer documentation file (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer', required=True, help='Manufacturer name')
    parser.add_argument('--model', required=True, help='Model name')
    parser.add_argument('-o', '--output', help='Output filename (default: auto-generated)')
    parser.add_argument('--protocol', default='modbusRTU', help='Protocol name (default: modbusRTU)')
    parser.add_argument('--category', default='Inverter', help='Device category (default: Inverter)')
    parser.add_argument('--sheet', help='Excel sheet name (processes all if not specified)')
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
    seen_tags = set()

    for df in dataframes:
        # Clean column names
        df.columns = [str(c).strip() for c in df.columns]

        # Identify columns with priority and prevent double-mapping
        col_map = {}
        used_cols = set()

        # Priority order: RegisterType > Address > Name > Type > Unit > Scale
        for key in ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Scale']:
            for col in df.columns:
                if col in used_cols:
                    continue

                col_lower = str(col).lower()
                for pattern in COLUMN_MAPPING[key]:
                    if pattern in col_lower:
                        col_map[key] = col
                        used_cols.add(col)
                        break
                if key in col_map:
                    break

        if 'Address' not in col_map and 'Name' not in col_map:
             logging.debug("Skipping table as neither Address nor Name columns found.")
             continue

        for _, row in df.iterrows():
            addr_raw = row.get(col_map.get('Address')) if 'Address' in col_map else None
            name_raw = row.get(col_map.get('Name')) if 'Name' in col_map else None

            if (addr_raw is None or pd.isna(addr_raw)) and (name_raw is None or pd.isna(name_raw)):
                continue

            # Skip header-like rows
            if str(addr_raw).lower() in COLUMN_MAPPING['Address'] or str(name_raw).lower() in COLUMN_MAPPING['Name']:
                continue

            addr = normalize_address(addr_raw)
            name = str(name_raw) if not pd.isna(name_raw) else f"Register {addr}"

            tag = generate_tag(name)
            # Ensure unique tag
            base_tag = tag if tag else "var"
            counter = 1
            while tag in seen_tags:
                tag = f"{base_tag}_{counter}"
                counter += 1
            seen_tags.add(tag)

            dtype_raw = row.get(col_map.get('Type')) if 'Type' in col_map else 'U16'
            unit_raw = row.get(col_map.get('Unit')) if 'Unit' in col_map else ''
            scale_raw = row.get(col_map.get('Scale')) if 'Scale' in col_map else '1'

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
                'Tag': tag,
                'RegisterType': row.get(col_map.get('RegisterType')) if 'RegisterType' in col_map else 'Holding Register',
                'Address': addr,
                'Type': normalize_type(dtype_raw),
                'Factor': scale if scale and scale != 'nan' else '1',
                'Offset': '0',
                'Unit': str(unit_raw) if not pd.isna(unit_raw) and str(unit_raw) != 'nan' else '',
                'Action': '4'
            }
            all_extracted_rows.append(extracted_row)

    if not all_extracted_rows:
        logging.error("No registers could be extracted from the tables.")
        sys.exit(1)

    logging.info(f"Successfully extracted {len(all_extracted_rows)} registers.")

    # Use DefFileGenerator logic to process and validate
    generator = Generator()
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
    for delimiter in [',', ';', '\t']:
        try:
            df = pd.read_csv(filepath, sep=delimiter, encoding='utf-8-sig')
            if len(df.columns) > 1:
                return [df]
        except Exception:
            continue
    return []

def load_excel(filepath, sheet_name=None):
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
                        df = pd.DataFrame(table[1:], columns=headers)
                        dfs.append(df)
    except Exception as e:
        logging.error(f"Error loading PDF file: {e}")
    return dfs

if __name__ == "__main__":
    main()
