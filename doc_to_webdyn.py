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

from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import run_generator

def main():
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', help='Path to documentation (PDF, Excel, CSV)')
    parser.add_argument('--manufacturer', required=True)
    parser.add_argument('--model', required=True)
    parser.add_argument('-o', '--output', help='Output filename')
    parser.add_argument('--protocol', default='modbusRTU')
    parser.add_argument('--category', default='Inverter')
    parser.add_argument('--sheet', help='Excel sheet')
    parser.add_argument('--pages', help='PDF pages (e.g. 1,2,5)')
    parser.add_argument('--address-offset', type=int, default=0)
    parser.add_argument('-v', '--verbose', action='store_true')

    args = parser.parse_args()

    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=log_level, format='%(levelname)s: %(message)s')

    extractor = Extractor()
    ext = os.path.splitext(args.input_file)[1].lower()

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xls']:
        raw_data = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = [int(p.strip()) for p in args.pages.split(',')] if args.pages else None
        raw_data = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv':
        raw_data = extractor.extract_from_csv(args.input_file)
    else:
        logging.error(f"Unsupported extension: {ext}")
        sys.exit(1)

    mapped_data = extractor.map_and_clean(raw_data)

    if not mapped_data:
        logging.error("No registers extracted.")
        sys.exit(1)

    import tempfile
    with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as tf:
        temp_csv = tf.name
        fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
        writer = csv.DictWriter(tf, fieldnames=fieldnames, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(mapped_data)

    try:
        # Determine output filename if not provided
        output_file = args.output
        if not output_file:
            safe_mfg = re.sub(r'[^a-zA-Z0-9]', '_', args.manufacturer).lower()
            safe_model = re.sub(r'[^a-zA-Z0-9]', '_', args.model).lower()
            output_file = f"{safe_mfg}_{safe_model}_definition.csv"

        run_generator(
            input_file=temp_csv,
            output=output_file,
            manufacturer=args.manufacturer,
            model=args.model,
            protocol=args.protocol,
            category=args.category,
            address_offset=args.address_offset
        )
    finally:
        if os.path.exists(temp_csv):
            os.remove(temp_csv)

if __name__ == "__main__":
    main()
