#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
import json
import tempfile
import csv
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator, run_generator, GeneratorConfig, generate_template

def setup_logging(verbose=False):
    logging.basicConfig(level=logging.DEBUG if verbose else logging.INFO,
                        format='%(levelname)s: %(message)s', force=True)

def main():
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', nargs='?', help='Path to documentation (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer', help='Manufacturer name')
    parser.add_argument('--model', help='Model name')
    parser.add_argument('-o', '--output', help='Output filename')
    parser.add_argument('--protocol', default='modbusRTU')
    parser.add_argument('--category', default='Inverter')
    parser.add_argument('--sheet', help='Excel sheet name')
    parser.add_argument('--pages', help='PDF pages (comma-separated)')
    parser.add_argument('--mapping', help='Mapping JSON file')
    parser.add_argument('--address-offset', type=int, default=0, help='Shift all addresses by offset')
    parser.add_argument('--forced-write', default='', help='Registers for forced write')
    parser.add_argument('--template', action='store_true', help='Generate a sample definition CSV')
    parser.add_argument('-v', '--verbose', action='store_true', help='Enable verbose logging')

    args = parser.parse_args()
    setup_logging(args.verbose)

    if args.template:
        out = args.output or "template_definition.csv"
        generate_template(out)
        logging.info(f"Template generated at {out}")
        return

    if not args.input_file or not args.manufacturer or not args.model:
        parser.print_help()
        sys.exit(1)

    mapping = {}
    if args.mapping:
        try:
            with open(args.mapping, 'r') as f:
                mapping = json.load(f)
        except Exception as e:
            logging.error(f"Error reading mapping JSON: {e}")

    ext = os.path.splitext(args.input_file)[1].lower()
    extractor = Extractor(mapping)

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        p_list = [int(p.strip()) for p in args.pages.split(',')] if args.pages else None
        raw = extractor.extract_from_pdf(args.input_file, p_list)
    elif ext == '.csv':
        raw = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml':
        raw = extractor.extract_from_xml(args.input_file)
    else:
        logging.error(f"Unsupported extension: {ext}")
        sys.exit(1)

    if not raw:
        logging.error("No data extracted.")
        sys.exit(1)

    # Apply offset once during mapping
    mapped = extractor.map_and_clean(raw, args.address_offset)
    if not mapped:
        logging.error("No registers extracted.")
        sys.exit(1)

    # Use run_generator logic via temp file to ensure consistency with main.py
    with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8', newline='') as tf:
        temp_csv = tf.name
        fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
        writer = csv.DictWriter(tf, fieldnames=fieldnames, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(mapped)

    try:
        output_file = args.output or f"{re.sub(r'[^a-zA-Z0-9]', '_', args.manufacturer).lower()}_{re.sub(r'[^a-zA-Z0-9]', '_', args.model).lower()}_definition.csv"
        config = GeneratorConfig(
            input_file=temp_csv,
            output=output_file,
            manufacturer=args.manufacturer,
            model=args.model,
            protocol=args.protocol,
            category=args.category,
            forced_write=args.forced_write,
            address_offset=0 # Already applied
        )
        run_generator(config)
    finally:
        if os.path.exists(temp_csv):
            os.remove(temp_csv)

if __name__ == "__main__":
    main()
