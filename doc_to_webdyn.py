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
    parser.add_argument('--manufacturer')
    parser.add_argument('--model')
    parser.add_argument('-o', '--output', help='Output filename')
    parser.add_argument('--protocol', default='modbusRTU')
    parser.add_argument('--category', default='Inverter')
    parser.add_argument('--sheet', help='Excel sheet name')
    parser.add_argument('--pages', help='PDF pages (comma-separated)')
    parser.add_argument('--address-offset', type=int, default=0)
    parser.add_argument('--forced-write', default='')
    parser.add_argument('--mapping', help='Mapping JSON')
    parser.add_argument('--template', action='store_true')
    parser.add_argument('-v', '--verbose', action='store_true')

    args = parser.parse_args()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format='%(levelname)s: %(message)s', force=True)

    if args.template:
        out = args.output or "template.csv"
        generate_template(out)
        logging.info(f"Template generated at {out}")
        return

    if not args.input_file or not args.manufacturer or not args.model:
        parser.error("input_file, --manufacturer, and --model are required unless --template is used.")

    mapping = {}
    if args.mapping:
        try:
            with open(args.mapping, 'r') as f:
                mapping = json.load(f)
        except Exception as e:
            logging.error(f"Error reading mapping JSON: {e}")

    ext_name = os.path.splitext(args.input_file)[1].lower()
    extractor = Extractor(mapping)

    if ext_name in ['.xlsx', '.xlsm', '.xltx', '.xltm']: raw = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext_name == '.pdf':
        pages = [int(p.strip()) for p in args.pages.split(',')] if args.pages else None
        raw = extractor.extract_from_pdf(args.input_file, pages)
    elif ext_name == '.csv': raw = extractor.extract_from_csv(args.input_file)
    elif ext_name == '.xml': raw = extractor.extract_from_xml(args.input_file)
    else: logging.error(f"Unsupported extension: {ext_name}"); sys.exit(1)

    if not raw: logging.error("No data extracted."); sys.exit(1)

    # Apply offset here to avoid double-offset in generator
    mapped = extractor.map_and_clean(raw, address_offset=args.address_offset)
    if not mapped: logging.error("No registers extracted."); sys.exit(1)

    generator = Generator()
    processed = generator.process_rows(mapped, address_offset=0) # Already applied in map_and_clean

    output_file = args.output or f"{re.sub(r'[^a-zA-Z0-9]', '_', args.manufacturer).lower()}_{re.sub(r'[^a-zA-Z0-9]', '_', args.model).lower()}_definition.csv"
    generator.write_output_csv(output_file, processed, args.manufacturer, args.model, args.protocol, args.category, args.forced_write)

if __name__ == "__main__":
    main()
