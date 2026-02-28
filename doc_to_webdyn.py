#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
import tempfile
import csv
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import run_generator
from DefFileGenerator.security_utils import ensure_safe_path

def main():
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', help='Path to the manufacturer documentation file (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer', required=True, help='Manufacturer name')
    parser.add_argument('--model', required=True, help='Model name')
    parser.add_argument('-o', '--output', help='Output filename (default: auto-generated)')
    parser.add_argument('--protocol', default='modbusRTU', help='Protocol name (default: modbusRTU)')
    parser.add_argument('--category', default='Inverter', help='Device category (default: Inverter)')
    parser.add_argument('--sheet', help='Excel sheet name (processes all if not specified)')
    parser.add_argument('--pages', help='PDF pages to extract (comma-separated, e.g. 1,2,5)')
    parser.add_argument('--address-offset', type=int, default=0, help='Value to subtract from register addresses')
    parser.add_argument('-v', '--verbose', action='store_true', help='Show detailed processing information')

    args = parser.parse_args()

    # Configure logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=log_level, format='%(levelname)s: %(message)s', force=True)

    logging.info(f"Processing {args.input_file} for {args.manufacturer} {args.model}")

    # Initialize Extractor
    extractor = Extractor()
    ext = os.path.splitext(args.input_file)[1].lower()
    all_tables = []

    if ext in ['.xlsx', '.xls', '.xlsm']:
        all_tables = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = [int(p.strip()) for p in args.pages.split(',')] if args.pages else None
        all_tables = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv':
        all_tables = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml':
        all_tables = extractor.extract_from_xml(args.input_file)
    else:
        logging.error(f"Unsupported file extension: {ext}")
        sys.exit(1)

    if not all_tables:
        logging.error("No data could be extracted from the file.")
        sys.exit(1)

    mapped_data = extractor.map_and_clean(all_tables)

    if not mapped_data:
        logging.error("No registers could be extracted from the tables.")
        sys.exit(1)

    # Determine output filename if not provided
    output_file = args.output
    if not output_file:
        safe_mfg = re.sub(r'[^a-zA-Z0-9]', '_', args.manufacturer).lower()
        safe_model = re.sub(r'[^a-zA-Z0-9]', '_', args.model).lower()
        output_file = f"{safe_mfg}_{safe_model}_definition.csv"

    # Use a temporary CSV for the intermediate mapped data
    with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as tf:
        temp_csv = tf.name
        fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
        writer = csv.DictWriter(tf, fieldnames=fieldnames, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(mapped_data)

    try:
        # Run generator
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
