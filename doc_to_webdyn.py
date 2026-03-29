#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
import json
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator, generate_template

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
    parser.add_argument('--forced-write', default='')
    parser.add_argument('--template', action='store_true', help='Generate a template definition file')
    parser.add_argument('--address-offset', type=int, default=0, help='Shift addresses by this value')
    parser.add_argument('-v', '--verbose', action='store_true')

    args = parser.parse_args()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format='%(levelname)s: %(message)s', force=True)

    if args.template:
        generate_template(args.output)
        return

    if not args.input_file or not args.manufacturer or not args.model:
        parser.print_help()
        logging.error("input_file, --manufacturer, and --model are required.")
        sys.exit(1)

    ext = os.path.splitext(args.input_file)[1].lower()
    mapping_data = {}
    if args.mapping:
        with open(args.mapping, 'r') as f:
            mapping_data = json.load(f)

    extractor = Extractor(mapping_data)

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = [int(p.strip()) for p in args.pages.split(',')] if args.pages else None
        raw = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv':
        raw = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml':
        raw = extractor.extract_from_xml(args.input_file)
    else:
        logging.error(f"Unsupported extension: {ext}")
        sys.exit(1)

    if not raw: logging.error("No data extracted."); sys.exit(1)

    # Apply offset during extraction
    mapped = extractor.map_and_clean(raw, address_offset=args.address_offset)
    if not mapped: logging.error("No registers extracted."); sys.exit(1)

    generator = Generator()
    # Offset is 0 here because it was applied in map_and_clean
    processed = generator.process_rows(mapped, address_offset=0)

    output_file = args.output or f"{re.sub(r'[^a-zA-Z0-9]', '_', args.manufacturer).lower()}_{re.sub(r'[^a-zA-Z0-9]', '_', args.model).lower()}_definition.csv"
    generator.write_output_csv(output_file, processed, args.manufacturer, args.model, args.protocol, args.category, args.forced_write)

if __name__ == "__main__":
    main()
