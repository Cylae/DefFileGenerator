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

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []
        logging.info(f"Extracting from Excel: {filepath}")
        wb = openpyxl.load_workbook(filepath, data_only=True)

        sheets = [sheet_name] if sheet_name else wb.sheetnames
        all_tables = []

        for name in sheets:
            if name not in wb.sheetnames:
                logging.warning(f"Sheet '{name}' not found.")
                continue
            ws = wb[name]
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
            target_pages = pdf.pages
            if pages:
                if isinstance(pages, int):
                    target_pages = [pdf.pages[pages-1]]
                elif isinstance(pages, list):
                    target_pages = [pdf.pages[i-1] for i in pages]

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
        # Manual delimiter detection: , ; or \t
        delimiters = [',', ';', '\t']
        best_delimiter = ','
        max_cols = 0

        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                sample = f.read(2048)
                for d in delimiters:
                    count = sample.count(d)
                    if count > max_cols:
                        max_cols = count
                        best_delimiter = d
        except Exception:
            pass

        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f, delimiter=best_delimiter)
                data = list(reader)
                return [data] if data else []
        except Exception as e:
            logging.error(f"Error extracting from CSV: {e}")
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
            logging.debug(f"Pandas XML extraction failed: {e}. Trying etree.")
            try:
                df = pd.read_xml(filepath, parser='etree')
                return [df.to_dict(orient='records')]
            except Exception as e2:
                logging.error(f"XML extraction failed: {e2}")
                return []

    def normalize_action(self, action):
        """Centralized Action normalization."""
        if not action: return '1'
        a = str(action).upper().strip()
        if a in ['R', 'READ']: return '4'
        if a in ['RW', 'W', 'WRITE']: return '1'
        return a

    def map_and_clean(self, tables):
        """Processes list of tables, mapping columns and cleaning data."""
        # Support single table input for backward compatibility
        if tables and isinstance(tables[0], dict):
            tables = [tables]

        generator = Generator()
        all_mapped_data = []

        for table in tables:
            if not table: continue

            first_row = table[0]
            col_map = {}
            used_src_cols = set()

            # 1. Apply explicit mapping from config
            for target, source in self.mapping.items():
                if source in first_row:
                    col_map[target] = source
                    used_src_cols.add(source)

            # 2. Priority-based heuristic mapping
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag', 'Factor', 'ScaleFactor']
            for target in detection_order:
                if target in col_map: continue
                for k in first_row.keys():
                    if k in used_src_cols: continue
                    k_lower = str(k).lower()
                    for pattern in self.COLUMN_MAPPING.get(target, []):
                        if pattern in k_lower:
                            col_map[target] = k
                            used_src_cols.add(k)
                            break
                    if target in col_map: break

            if 'Address' not in col_map and 'Name' not in col_map:
                continue

            for row in table:
                new_row = {}
                for target, source in col_map.items():
                    val = row.get(source)
                    # Support fraction parsing for Factor (e.g. "1/10")
                    if target == 'Factor' and val and '/' in str(val):
                        try:
                            parts = str(val).split('/')
                            val = str(float(parts[0]) / float(parts[1]))
                        except (ValueError, ZeroDivisionError):
                            pass
                    new_row[target] = val

                # Passthrough unmapped columns
                for k, v in row.items():
                    if k not in used_src_cols:
                        new_row[k] = v

                # Basic validation: must have Name or Address
                if not new_row.get('Name') and not new_row.get('Address'):
                    continue

                # Normalization is mostly delegated to Generator.process_rows,
                # but we handle some initial cleaning here for the intermediate CSV.

                # Normalize Address
                if new_row.get('Address'):
                    addr = str(new_row['Address']).strip()
                    # Handle Address_Length or Address_Start_Bit patterns
                    if '_' in addr:
                        parts = addr.split('_')
                        norm_parts = [generator.normalize_address_val(p) for p in parts]
                        new_row['Address'] = '_'.join(norm_parts)
                    else:
                        new_row['Address'] = generator.normalize_address_val(addr)

                # Normalize Action if present
                if new_row.get('Action'):
                    new_row['Action'] = self.normalize_action(new_row['Action'])

                all_mapped_data.append(new_row)

        return all_mapped_data

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

    parser = argparse.ArgumentParser(description='Extract register information from various files.')
    parser.add_argument('input_file', help='Path to the source file (PDF, Excel, CSV, XML).')
    parser.add_argument('-o', '--output', help='Path to the output CSV file.')
    parser.add_argument('--mapping', help='JSON file containing column mapping.')
    parser.add_argument('--sheet', help='Excel sheet name to extract from.')
    parser.add_argument('--pages', help='PDF pages to extract from (e.g. 1,2,5).')

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

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw_data = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = [int(p.strip()) for p in args.pages.split(',')] if args.pages else None
        raw_data = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv':
        raw_data = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml':
        raw_data = extractor.extract_from_xml(args.input_file)
    else:
        logging.error(f"Unsupported extension: {ext}")
        sys.exit(1)

    if not raw_data:
        logging.error("No data extracted.")
        sys.exit(1)

    mapped_data = extractor.map_and_clean(raw_data)

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
