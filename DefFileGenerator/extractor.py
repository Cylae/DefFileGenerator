#!/usr/bin/env python3
import argparse
import csv
import json
import logging
import os
import re
import sys
import io

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
    from defusedxml import ElementTree as ET
    HAS_DEFUSEDXML = True
except ImportError:
    HAS_DEFUSEDXML = False

# Fallback for relative vs absolute import depending on how it's executed
try:
    from DefFileGenerator.def_gen import Generator
except ImportError:
    from def_gen import Generator

class Extractor:
    COLUMN_MAPPING = {
        'RegisterType': ['register type', 'reg type', 'modbus type', 'registertype', 'function'],
        'Address': ['address', 'addr', 'offset', 'register', 'reg', 'index'],
        'Name': ['name', 'description', 'parameter', 'variable', 'signal', 'signal name', 'point'],
        'Type': ['data type', 'datatype', 'type', 'format', 'data format'],
        'Unit': ['unit', 'units'],
        'Tag': ['tag'],
        'Action': ['action', 'access', 'read/write', 'r/w'],
        'Factor': ['scale', 'factor', 'multiplier', 'ratio', 'coefficient a', 'multiplier factor'],
        'Offset': ['offset', 'bias', 'coefficient b'],
        'ScaleFactor': ['scalefactor', 'scale factor', 'exponent'],
        'StartBit': ['startbit', 'bit offset', 'bit', 'start'],
        'Length': ['length', 'len', 'size', 'count', 'quantity']
    }

    def __init__(self, mapping=None):
        self.mapping = mapping or {}

    def extract_from_excel(self, filepath, sheet_name=None):
        """Extracts data from Excel sheets. Returns a list of tables (list of lists of dicts)."""
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []

        tables = []
        wb = openpyxl.load_workbook(filepath, data_only=True)

        sheets = [wb[sheet_name]] if sheet_name else wb.worksheets

        for ws in sheets:
            data = []
            rows = list(ws.rows)
            if not rows: continue

            # Find headers (look in first few rows)
            header_row_idx = 0
            headers = []
            for i in range(min(5, len(rows))):
                row_vals = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[i]]
                if any(any(p in val.lower() for patterns in self.COLUMN_MAPPING.values() for p in patterns) for val in row_vals):
                    headers = row_vals
                    header_row_idx = i
                    break

            if not headers:
                headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]
                header_row_idx = 0

            for row in rows[header_row_idx+1:]:
                row_dict = {}
                for i, cell in enumerate(row):
                    if i < len(headers) and headers[i]:
                        row_dict[headers[i]] = cell.value
                if any(v is not None for v in row_dict.values()):
                    data.append(row_dict)
            if data:
                tables.append(data)
        return tables

    def extract_from_pdf(self, filepath, pages=None):
        """Extracts tables from PDF pages. Returns a list of tables."""
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []
        all_tables = []
        try:
            with pdfplumber.open(filepath) as pdf:
                target_pages = pdf.pages if pages is None else [pdf.pages[i-1] for i in (pages if isinstance(pages, list) else [pages])]
                for page in target_pages:
                    tables = page.extract_tables()
                    logging.debug(f"Found {len(tables)} tables on page {page.page_number}")
                    for table in tables:
                        if not table or len(table) < 2: continue
                        headers = [str(c).replace('\n', ' ').strip() if c else "" for c in table[0]]
                        table_data = []
                        for row in table[1:]:
                            row_dict = {}
                            for i, cell in enumerate(row):
                                if i < len(headers) and headers[i]:
                                    row_dict[headers[i]] = str(cell).replace('\n', ' ').strip() if cell else ""
                            if any(v for v in row_dict.values()):
                                table_data.append(row_dict)
                        if table_data:
                            all_tables.append(table_data)
        except Exception as e:
            logging.error(f"Error extracting from PDF {filepath}: {e}")
        return all_tables

    def extract_from_csv(self, filepath):
        """Extracts data from CSV. Returns a list containing one table."""
        try:
            with open(filepath, 'rb') as f:
                content = f.read()
                encoding = 'utf-16' if content.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'
            with open(filepath, 'r', encoding=encoding) as f:
                snippet = f.read(2048)
                f.seek(0)
                delimiter = ','
                for d in [',', ';', '\t']:
                    if d in snippet:
                        delimiter = d; break
                reader = csv.DictReader(f, delimiter=delimiter)
                return [list(reader)]
        except Exception as e:
            logging.error(f"Error extracting from CSV: {e}")
            return []

    def extract_from_xml(self, filepath):
        """Extracts data from XML securely without pandas. Returns a list of tables."""
        if not HAS_DEFUSEDXML:
            logging.error("defusedxml is required for secure XML parsing.")
            return []
        try:
            tree = ET.parse(filepath)
            root = tree.getroot()

            # Simple heuristic: find elements that have multiple children with same name
            # or just flatten the structure.
            # Most register XMLs are flat lists of elements.
            data = []
            for child in root:
                row = {}
                for subchild in child:
                    row[subchild.tag] = subchild.text
                if row:
                    data.append(row)
            return [data] if data else []
        except Exception as e:
            logging.error(f"Error extracting from XML: {e}")
            return []

    def map_and_clean(self, tables):
        """Maps manufacturer columns to standard ones and cleans data."""
        if not tables: return []
        # Ensure tables is a list of lists of dicts
        if isinstance(tables, list) and tables and isinstance(tables[0], dict):
            tables = [tables]

        final_data = []
        generator = Generator()

        for table in tables:
            if not table: continue

            # Use first 5 rows to detect columns to be more robust
            col_map = {}
            used_src_cols = set()

            # 1. Explicit mapping from user
            for target, source in self.mapping.items():
                for row in table[:5]:
                    if source in row:
                        col_map[target] = source
                        used_src_cols.add(source)
                        break

            # 2. Heuristic fuzzy matching
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag', 'Factor', 'Offset', 'ScaleFactor', 'StartBit', 'Length']
            all_src_cols = set()
            for row in table[:5]:
                all_src_cols.update(row.keys())

            for target in detection_order:
                if target in col_map: continue
                patterns = self.COLUMN_MAPPING.get(target, [target.lower()])
                for src_col in all_src_cols:
                    if src_col in used_src_cols: continue
                    if any(p in str(src_col).lower() for p in patterns):
                        col_map[target] = src_col
                        used_src_cols.add(src_col)
                        break

            for row in table:
                new_row = {target: row.get(src_col) for target, src_col in col_map.items()}

                # Check for Address vs Offset ambiguity (often 'offset' is used for address)
                if not new_row.get('Address') and new_row.get('Offset'):
                    # If we found an 'Offset' column but no 'Address', it's likely the address
                    new_row['Address'] = new_row['Offset']
                    new_row['Offset'] = '0'

                if not new_row.get('Name') and not new_row.get('Address'): continue

                # Compound address construction (Addr_StartBit_Length)
                addr = str(new_row.get('Address', '')).strip()
                sbit = str(new_row.get('StartBit', '')).strip() if new_row.get('StartBit') is not None else ""
                length = str(new_row.get('Length', '')).strip() if new_row.get('Length') is not None else ""

                if sbit:
                    # If we have a start bit, it's likely a BITS type if not specified
                    if not new_row.get('Type'):
                        new_row['Type'] = 'BITS'
                    if not length: length = '1'
                    addr = f"{addr}_{sbit}_{length}"
                elif length and Generator.normalize_type(new_row.get('Type')) == 'STRING':
                    addr = f"{addr}_{length}"

                # Normalize Address via Generator
                if '_' in addr:
                    new_row['Address'] = '_'.join(Generator.normalize_address_val(p) for p in addr.split('_'))
                else:
                    new_row['Address'] = Generator.normalize_address_val(addr)

                # Normalize Type via Generator
                new_row['Type'] = Generator.normalize_type(new_row.get('Type', 'U16'))

                # Normalize Factor
                if new_row.get('Factor'):
                    new_row['Factor'] = str(Generator._parse_numeric(new_row['Factor'], 1.0))

                if 'RegisterType' not in new_row:
                    new_row['RegisterType'] = 'Holding Register'

                final_data.append(new_row)
        return final_data

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(description='Extract register information.')
    parser.add_argument('input_file'); parser.add_argument('-o', '--output')
    parser.add_argument('--mapping'); parser.add_argument('--sheet'); parser.add_argument('--pages')
    parser.add_argument('--address-offset', type=int, default=0)
    args = parser.parse_args()

    mapping = {}
    if args.mapping:
        with open(args.mapping, 'r') as f: mapping = json.load(f)

    extractor = Extractor(mapping)
    ext = os.path.splitext(args.input_file)[1].lower()

    if ext in ['.xlsx', '.xlsm']: raw = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf': raw = extractor.extract_from_pdf(args.input_file, [int(p) for p in args.pages.split(',')] if args.pages else None)
    elif ext == '.csv': raw = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml': raw = extractor.extract_from_xml(args.input_file)
    else: logging.error(f"Unsupported extension: {ext}"); sys.exit(1)

    mapped = extractor.map_and_clean(raw)

    # Apply address offset if specified in extraction standalone
    if args.address_offset != 0:
        generator = Generator()
        for row in mapped:
            row['Address'] = generator.apply_address_offset(row['Address'], args.address_offset)

    out = open(args.output, 'w', newline='', encoding='utf-8') if args.output else sys.stdout
    fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
    writer = csv.DictWriter(out, fieldnames=fieldnames, extrasaction='ignore')
    writer.writeheader(); writer.writerows(mapped)
    if args.output: out.close()

if __name__ == "__main__":
    main()
