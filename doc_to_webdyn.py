#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
import json
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator, generate_template

RE_TAG_CLEAN = re.compile(r'[^a-zA-Z0-9]')

def main():
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', nargs='?', help='Path to documentation (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer', help='Manufacturer name')
    parser.add_argument('--model', help='Model name')
    parser.add_argument('-o', '--output', help='Output filename')
    parser.add_argument('--protocol', default='modbusRTU')
    parser.add_argument('--category', default='Inverter')
    parser.add_argument('--sheet', help='Excel sheet name')
    parser.add_argument('--address-offset', type=int, default=0, help='Address offset')
    parser.add_argument('--forced-write', default='', help='Forced write value')
    parser.add_argument('--mapping', help='JSON mapping file')
    parser.add_argument('--pages', help='Comma-separated list of pages for PDF')
    parser.add_argument('--template', action='store_true', help='Generate a template definition CSV')
    parser.add_argument('-v', '--verbose', action='store_true')

    args = parser.parse_args()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format='%(levelname)s: %(message)s', force=True)

    if args.template:
        generate_template(args.output)
        return

    if not args.input_file or not args.manufacturer or not args.model:
        parser.error("the following arguments are required: input_file, --manufacturer, --model (unless using --template)")

    mapping = {}
    if args.mapping:
        try:
            with open(args.mapping, 'r') as f:
                mapping = json.load(f)
        except Exception as e:
            logging.error(f"Error loading mapping file: {e}")
            sys.exit(1)

    pages = None
    if args.pages:
        try:
            pages = [int(p.strip()) for p in args.pages.split(',')]
        except ValueError:
            logging.error("Invalid --pages argument. Must be comma-separated integers.")
            sys.exit(1)

    ext = os.path.splitext(args.input_file)[1].lower()
    extractor = Extractor(mapping)

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        raw = extractor.extract_from_pdf(args.input_file, pages=pages)
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

    mapped = extractor.map_and_clean(raw)
    if not mapped:
        logging.error("No registers extracted.")
        sys.exit(1)

    generator = Generator()
    processed = generator.process_rows(mapped, address_offset=args.address_offset)

    output_file = args.output or f"{RE_TAG_CLEAN.sub('_', args.manufacturer).lower()}_{RE_TAG_CLEAN.sub('_', args.model).lower()}_definition.csv"
    generator.write_output_csv(output_file, processed, args.manufacturer, args.model, args.protocol, args.category, args.forced_write)

if __name__ == "__main__":
    main()
