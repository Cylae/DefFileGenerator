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

try:
    from DefFileGenerator.def_gen import Generator
except ImportError:
    from def_gen import Generator

class Extractor:
    COLUMN_MAPPING = {
        'RegisterType': ['register type', 'reg type', 'modbus type', 'registertype'],
        'Address': ['address', 'addr', 'offset', 'register', 'reg'],
        'Name': ['name', 'description', 'parameter', 'variable', 'signal', 'signal name'],
        'Type': ['data type', 'datatype', 'type', 'format'],
        'Unit': ['unit', 'units'],
        'Tag': ['tag'],
        'Action': ['action', 'access'],
        'Factor': ['scale', 'factor', 'multiplier', 'ratio'],
        'ScaleFactor': ['scalefactor'],
        'Offset': ['offset', 'bias', 'coefficient b']
    }

    def __init__(self, mapping=None):
        self.mapping = mapping or {}

    def extract_from_excel(self, filepath, sheet_name=None):
        """Extracts all tables from an Excel file. Returns list of lists of dicts."""
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []

        all_tables = []
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            sheets = [wb[sheet_name]] if sheet_name else wb.worksheets

            for ws in sheets:
                data = []
                rows = list(ws.rows)
                if not rows: continue

                headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]
                for row in rows[1:]:
                    row_dict = {headers[i]: cell.value for i, cell in enumerate(row) if i < len(headers)}
                    if any(row_dict.values()):
                        data.append(row_dict)
                if data:
                    all_tables.append(data)
        except Exception as e:
            logging.error(f"Error extracting from Excel {filepath}: {e}")

        return all_tables

    def extract_from_pdf(self, filepath, pages=None):
        """Extracts all tables from a PDF file. Returns list of lists of dicts."""
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
                        data = []
                        for row in table[1:]:
                            row_dict = {}
                            for i, cell in enumerate(row):
                                if i < len(headers):
                                    row_dict[headers[i]] = str(cell).replace('\n', ' ').strip() if cell else ""
                            if any(row_dict.values()):
                                data.append(row_dict)
                        if data:
                            all_tables.append(data)
        except Exception as e:
            logging.error(f"Error extracting from PDF {filepath}: {e}")
        return all_tables

    def extract_from_csv(self, filepath):
        """Extracts table from a CSV file. Returns list of lists of dicts."""
        try:
            with open(filepath, 'rb') as f:
                content = f.read()
                encoding = 'utf-16' if content.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'
            with open(filepath, 'r', encoding=encoding) as f:
                snippet = f.read(1024)
                f.seek(0)
                delimiter = ','
                for d in [',', ';', '\t']:
                    if d in snippet:
                        delimiter = d; break
                reader = csv.DictReader(f, delimiter=delimiter)
                data = list(reader)
                return [data] if data else []
        except Exception as e:
            logging.error(f"Error extracting from CSV: {e}")
            return []

    def extract_from_xml(self, filepath):
        """Extracts data from XML securely using defusedxml. Returns list of lists of dicts."""
        if not HAS_DEFUSEDXML:
            logging.error("defusedxml is required for secure XML parsing.")
            return []
        try:
            with open(filepath, 'rb') as f:
                tree = ET.parse(f)
                root = tree.getroot()

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
        """Maps manufacturer columns to standard columns and cleans data."""
        if not tables: return []
        # Support single table (list of dicts) or list of tables
        if isinstance(tables, list) and tables and isinstance(tables[0], dict):
            tables = [tables]

        final_data = []

        for table in tables:
            if not table: continue

            # 1. Identify source columns by scanning first 5 rows
            all_src_cols = set()
            for row in table[:5]:
                all_src_cols.update(row.keys())

            col_map = {}
            used_src_cols = set()

            # 2. Explicit mapping from user
            for target, source in self.mapping.items():
                if source in all_src_cols:
                    col_map[target] = source
                    used_src_cols.add(source)

            # 3. Priority fuzzy matching
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag', 'Factor', 'ScaleFactor', 'Offset']
            for target in detection_order:
                if target in col_map: continue
                patterns = self.COLUMN_MAPPING.get(target, [target.lower()])
                # Prioritize exact matches first
                found = False
                for src_col in all_src_cols:
                    if src_col in used_src_cols: continue
                    if str(src_col).lower().strip() in patterns:
                        col_map[target] = src_col
                        used_src_cols.add(src_col)
                        found = True
                        break
                if found: continue

                # Then fuzzy matches
                for src_col in all_src_cols:
                    if src_col in used_src_cols: continue
                    if any(p in str(src_col).lower() for p in patterns):
                        col_map[target] = src_col
                        used_src_cols.add(src_col)
                        break

            # 4. Process rows using identified column mapping
            for row in table:
                new_row = {target: row.get(src_col) for target, src_col in col_map.items()}
                if not new_row.get('Name') and not new_row.get('Address'): continue

                # Normalize Address using Generator's logic
                addr = str(new_row.get('Address', '')).strip()
                if addr:
                    if '_' in addr:
                        parts = addr.split('_')
                        new_row['Address'] = '_'.join(Generator.normalize_address_val(p) for p in parts)
                    else:
                        new_row['Address'] = Generator.normalize_address_val(addr)

                # Normalize Type using Generator's logic
                new_row['Type'] = Generator.normalize_type(new_row.get('Type', 'U16'))

                # Normalize Factor (fractions like 1/10)
                factor_val = new_row.get('Factor')
                if factor_val and '/' in str(factor_val):
                    try:
                        p = str(factor_val).split('/')
                        new_row['Factor'] = str(float(p[0]) / float(p[1]))
                    except (ValueError, ZeroDivisionError):
                        pass # Let Generator._parse_numeric handle it later

                if 'RegisterType' not in new_row or not new_row['RegisterType']:
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

    # Apply address offset if requested during extraction phase (optional, usually done in generation)
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
