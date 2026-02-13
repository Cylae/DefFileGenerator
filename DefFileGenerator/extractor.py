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
    def __init__(self, mapping=None):
        self.mapping = mapping or {}
        self.generator = Generator()

    def extract_from_excel(self, filepath, sheet_name=None):
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return []

        logging.info(f"Extracting from Excel: {filepath}")
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            ws = wb[sheet_name] if sheet_name else wb.active

            data = []
            rows = list(ws.rows)
            if not rows: return []

            headers = [str(cell.value).strip() if cell.value is not None else "" for cell in rows[0]]
            for row in rows[1:]:
                row_data = {}
                for i, cell in enumerate(row):
                    if i < len(headers):
                        row_data[headers[i]] = cell.value
                data.append(row_data)
            return data
        except Exception as e:
            logging.error(f"Excel extraction failed: {e}")
            return []

    def extract_from_pdf(self, filepath, pages=None):
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return []

        logging.info(f"Extracting from PDF: {filepath}")
        data = []
        try:
            with pdfplumber.open(filepath) as pdf:
                if pages:
                    target_pages = [pdf.pages[i-1] for i in pages if 0 < i <= len(pdf.pages)]
                else:
                    target_pages = pdf.pages

                for page in target_pages:
                    for table in page.extract_tables() or []:
                        if len(table) < 2: continue
                        headers = [str(c).replace('\n', ' ').strip() if c else "" for c in table[0]]
                        for row in table[1:]:
                            row_data = {}
                            for i, cell in enumerate(row):
                                if i < len(headers):
                                    row_data[headers[i]] = str(cell).replace('\n', ' ').strip() if cell else ""
                            data.append(row_data)
            return data
        except Exception as e:
            logging.error(f"PDF extraction failed: {e}")
            return []

    def map_and_clean(self, raw_data):
        if not raw_data: return []

        # Identify mapping
        first_row = raw_data[0]
        col_map = {}
        assigned = set()

        # Explicit mapping
        for target, source in self.mapping.items():
            if source in first_row:
                col_map[target] = source
                assigned.add(source)

        # Fuzzy matching for standard columns
        standard = ['RegisterType', 'Name', 'Address', 'Type', 'Unit', 'Tag', 'Action', 'Factor', 'ScaleFactor', 'Offset']
        for target in standard:
            if target in col_map: continue
            for k in first_row.keys():
                if k in assigned: continue
                if k.lower() == target.lower() or target.lower() in k.lower():
                    col_map[target] = k
                    assigned.add(k)
                    break

        mapped_data = []
        for row in raw_data:
            new_row = {}
            for target, source in col_map.items():
                new_row[target] = row.get(source)

            # Additional unmapped columns
            for k, v in row.items():
                if k not in assigned and k not in new_row:
                    new_row[k] = v

            # Initial normalization
            if new_row.get('Address'):
                a = str(new_row['Address']).strip()
                if '_' in a:
                    new_row['Address'] = '_'.join([self.generator.normalize_address_val(p) for p in a.split('_')])
                else:
                    new_row['Address'] = self.generator.normalize_address_val(a)

            if new_row.get('Type'):
                new_row['Type'] = self.generator.normalize_type(new_row['Type'])

            if new_row.get('Action'):
                new_row['Action'] = self.generator.normalize_action(new_row['Action'])

            if new_row.get('Name'):
                mapped_data.append(new_row)

        return mapped_data

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser()
    parser.add_argument('input_file')
    parser.add_argument('-o', '--output')
    parser.add_argument('--mapping')
    parser.add_argument('--sheet')
    parser.add_argument('--pages')
    args = parser.parse_args()

    mapping = {}
    if args.mapping:
        try:
            with open(args.mapping, 'r') as f: mapping = json.load(f)
        except Exception as e:
            logging.error(f"Mapping load failed: {e}")
            sys.exit(1)

    extractor = Extractor(mapping)
    ext = os.path.splitext(args.input_file)[1].lower()

    if ext in ['.xlsx', '.xls']: raw_data = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = [int(p) for p in args.pages.split(',')] if args.pages else None
        raw_data = extractor.extract_from_pdf(args.input_file, pages)
    else:
        logging.error(f"Unsupported extension: {ext}")
        sys.exit(1)

    mapped = extractor.map_and_clean(raw_data)
    if not mapped:
        logging.error("No data extracted.")
        sys.exit(1)

    out = open(args.output, 'w', newline='', encoding='utf-8') if args.output else sys.stdout
    fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
    writer = csv.DictWriter(out, fieldnames=fieldnames, extrasaction='ignore')
    writer.writeheader()
    writer.writerows(mapped)
    if args.output: out.close()
