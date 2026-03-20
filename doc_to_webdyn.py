#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
import json
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator, generate_template

# Pre-compiled regex for cleaning manufacturer/model names in output filename
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
    parser.add_argument('--pages', help='PDF pages (comma-separated integers)')
    parser.add_argument('--forced-write', default='', help='Forced write value')
    parser.add_argument('--address-offset', type=int, default=0, help='Global address offset')
    parser.add_argument('--mapping', help='Path to mapping JSON file')
    parser.add_argument('--template', action='store_true', help='Generate a template input CSV')
    parser.add_argument('-v', '--verbose', action='store_true')

    args = parser.parse_args()

    # Handle template generation separately
    if args.template:
        generate_template(args.output)
        return

    # Basic validation for non-template mode
    if not args.input_file or not args.manufacturer or not args.model:
        if not args.input_file:
             parser.error("the following arguments are required: input_file")
        if not args.manufacturer:
             parser.error("the following arguments are required: --manufacturer")
        if not args.model:
             parser.error("the following arguments are required: --model")

    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format='%(levelname)s: %(message)s',
                        force=True)

    mapping = {}
    if args.mapping:
        try:
            with open(args.mapping, 'r') as f:
                mapping = json.load(f)
        except Exception as e:
            logging.error(f"Error loading mapping JSON: {e}")
            sys.exit(1)

    ext = os.path.splitext(args.input_file)[1].lower()
    extractor = Extractor(mapping)

    # Process pages argument if provided
    pages = None
    if args.pages:
        try:
            pages = [int(p.strip()) for p in args.pages.split(',')]
        except ValueError:
            logging.error("Invalid --pages argument. Must be comma-separated integers.")
            sys.exit(1)

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        raw = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv':
        raw = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml':
        raw = extractor.extract_from_xml(args.input_file)
    else:
        logging.error(f"Unsupported extension: {ext}"); sys.exit(1)

    if not raw:
        logging.error("No data extracted."); sys.exit(1)

    # Apply address_offset during extraction, pass 0 to process_rows later
    mapped = extractor.map_and_clean(raw, address_offset=args.address_offset)
    if not mapped:
        logging.error("No registers extracted."); sys.exit(1)

    generator = Generator()
    # Pass address_offset=0 here because it was already applied in map_and_clean
    processed = generator.process_rows(mapped, address_offset=0)

    if not args.output:
        m_clean = RE_TAG_CLEAN.sub('_', args.manufacturer).lower()
        mo_clean = RE_TAG_CLEAN.sub('_', args.model).lower()
        output_file = f"{m_clean}_{mo_clean}_definition.csv"
    else:
        output_file = args.output

    generator.write_output_csv(output_file, processed, args.manufacturer, args.model,
                               args.protocol, args.category, args.forced_write)

if __name__ == "__main__":
    main()
