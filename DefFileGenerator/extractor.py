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

try:
    from defusedxml import ElementTree as ET
    HAS_DEFUSEDXML = True
except ImportError:
    HAS_DEFUSEDXML = False

import io
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
        'Tag': ['tag'],
        'Offset': ['offset_val'],
        'ScaleFactor': ['scalefactor']
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
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                logging.error(f"Sheet '{sheet_name}' not found in {filepath}")
                return []
            ws = wb[sheet_name]
        else:
            ws = wb.active

        data = []
        rows = list(ws.rows)
        if not rows:
            return []

        headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]

        for row_idx, row in enumerate(rows[1:], start=2):
            row_data = {}
            for i, cell in enumerate(row):
                if i < len(headers):
                    row_data[headers[i]] = cell.value
            data.append(row_data)
        return data

    def extract_from_csv(self, filepath):
        logging.info(f"Extracting from CSV: {filepath}")
        data = []
        # Detect encoding and delimiter
        encoding = 'utf-8-sig'
        try:
            with open(filepath, 'rb') as f:
                header = f.read(2048)
                if b'\xff\xfe' in header or b'\xfe\xff' in header:
                    encoding = 'utf-16'
        except Exception:
            pass

        for delimiter in [',', ';', '\t']:
            try:
                with open(filepath, 'r', encoding=encoding) as f:
                    reader = csv.DictReader(f, delimiter=delimiter)
                    rows = list(reader)
                    if rows and len(reader.fieldnames) > 1:
                        # Clean fieldnames
                        reader.fieldnames = [fn.strip() for fn in reader.fieldnames]
                        # Re-read with cleaned fieldnames if necessary or just use the rows
                        data = rows
                        break
            except Exception:
                continue
        return data

    def extract_from_xml(self, filepath):
        if not HAS_PANDAS:
            logging.error("pandas is required for XML extraction.")
            return []
        if not HAS_DEFUSEDXML:
            logging.error("defusedxml is required for secure XML parsing. Aborting.")
            return []

        logging.info(f"Extracting from XML: {filepath}")
        try:
            with open(filepath, 'rb') as f:
                xml_data = f.read()
            # Validate with defusedxml
            ET.fromstring(xml_data)
            # Use pandas to read
            df = pd.read_xml(io.BytesIO(xml_data))
            return df.to_dict(orient='records')
        except Exception as e:
            logging.error(f"Error loading XML file: {e}")
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
        mapped_data = []
        # Mapping: target_col -> source_col
        # Default target cols: Name, Tag, RegisterType, Address, Type, Factor, Offset, Unit, Action, ScaleFactor

        # Identify standard columns once to avoid repeated fuzzy matching
        first_row = raw_data[0] if raw_data else {}
        standard_cols_mapping = {}
        assigned_keys = set()

        # Explicitly mapped columns from config
        for target, source in self.mapping.items():
            if source in first_row:
                standard_cols_mapping[target] = source
                assigned_keys.add(source)

        # Fuzzy match for standard columns if not explicitly mapped
        # Priority order to avoid misidentification (e.g. RegisterType as Type)
        detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Factor', 'Action', 'Tag']
        for target in detection_order:
            if target not in standard_cols_mapping:
                for k in first_row.keys():
                    if k in assigned_keys:
                        continue
                    k_lower = k.lower()
                    for pattern in self.COLUMN_MAPPING.get(target, []):
                        if pattern in k_lower:
                            standard_cols_mapping[target] = k
                            assigned_keys.add(k)
                            break
                    if target in standard_cols_mapping:
                        break

        generator = Generator()
        for row in raw_data:
            new_row = {}
            for target, source in standard_cols_mapping.items():
                if source in row:
                    new_row[target] = row[source]

            # Fill in other columns that might not be in standard_cols but are in row
            for k, v in row.items():
                if k not in assigned_keys and k not in new_row:
                    new_row[k] = v

            # Clean Address using Generator's logic
            if 'Address' in new_row and new_row['Address']:
                addr = str(new_row['Address']).strip()
                # Support Address_Length and Address_Start_Bit formats (e.g. 30001_10 or 30001_0_1)
                match = re.search(r'(0x[0-9A-Fa-f]+|\d+)(_(\d+))+', addr)
                if match:
                    addr = match.group(0)

                if '_' in addr:
                    parts = addr.split('_')
                    norm_parts = [generator.normalize_address_val(p) for p in parts]
                    new_row['Address'] = '_'.join(norm_parts)
                else:
                    new_row['Address'] = generator.normalize_address_val(addr)

            # Clean Factor (fraction parsing)
            if 'Factor' in new_row and new_row['Factor']:
                factor = str(new_row['Factor']).strip()
                if '/' in factor:
                    try:
                        parts = factor.split('/')
                        new_row['Factor'] = str(float(parts[0]) / float(parts[1]))
                    except (ValueError, ZeroDivisionError):
                        pass

            # Clean Type
            if 'Type' in new_row:
                new_row['Type'] = self.normalize_type(new_row['Type'])

            # Ensure mandatory fields for def_gen
            if 'Name' not in new_row or not new_row['Name']:
                continue # Skip rows without a name

            # Tag and RegisterType will be handled by Generator if missing,
            # but we can set default RegisterType here for the simplified CSV.
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
