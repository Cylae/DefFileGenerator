#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
import json
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator, RE_TAG_CLEAN

def main():
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', nargs='?', help='Path to documentation (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer')
    parser.add_argument('--model')
    parser.add_argument('-o', '--output', help='Output filename')
    parser.add_argument('--protocol', default='modbusRTU')
    parser.add_argument('--category', default='Inverter')
    parser.add_argument('--sheet', help='Excel sheet name')
    parser.add_argument('--pages', help='PDF pages (comma-separated list)')
    parser.add_argument('--address-offset', type=int, default=0)
    parser.add_argument('--mapping', help='Mapping JSON file')
    parser.add_argument('--forced-write', default='')
    parser.add_argument('--template', action='store_true', help='Generate a sample template definition CSV')
    parser.add_argument('-v', '--verbose', action='store_true')

    args = parser.parse_args()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format='%(levelname)s: %(message)s', force=True)

    generator = Generator()

    if args.template:
        generator.generate_template(args.output)
        return

    if not args.input_file or not args.manufacturer or not args.model:
        parser.error("input_file, --manufacturer, and --model are required unless --template is used.")

    mapping = {}
    if args.mapping:
        try:
            with open(args.mapping, 'r') as f:
                mapping = json.load(f)
        except Exception as e:
            logging.error(f"Error reading mapping file: {e}")
            sys.exit(1)

    ext = os.path.splitext(args.input_file)[1].lower()
    extractor = Extractor(mapping)

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        try:
            pages = [int(p.strip()) for p in args.pages.split(',')] if args.pages else None
        except ValueError:
            logging.error("Invalid pages format. Use comma-separated integers.")
            sys.exit(1)
        raw = extractor.extract_from_pdf(args.input_file, pages)
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

    processed = generator.process_rows(mapped, args.address_offset)

    if args.output:
        output_file = args.output
    else:
        m_tag = RE_TAG_CLEAN.sub('_', args.manufacturer).lower()
        mod_tag = RE_TAG_CLEAN.sub('_', args.model).lower()
        output_file = f"{m_tag}_{mod_tag}_definition.csv"

    generator.write_output_csv(output_file, processed, args.manufacturer, args.model,
                               args.protocol, args.category, args.forced_write)

if __name__ == "__main__":
    main()
