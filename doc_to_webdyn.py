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
    parser.add_argument('input_file', help='Path to documentation (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer', required=True)
    parser.add_argument('--model', required=True)
    parser.add_argument('-o', '--output', help='Output filename')
    parser.add_argument('--protocol', default='modbusRTU')
    parser.add_argument('--category', default='Inverter')
    parser.add_argument('--sheet', help='Excel sheet name')
    parser.add_argument('--pages', help='Comma-separated list of PDF pages')
    parser.add_argument('--mapping', help='JSON mapping file')
    parser.add_argument('--forced-write', default='')
    parser.add_argument('--address-offset', type=int, default=0)
    parser.add_argument('-v', '--verbose', action='store_true')

    args = parser.parse_args()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format='%(levelname)s: %(message)s', force=True)

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

    if not raw:
        logging.error("No data extracted.")
        sys.exit(1)

    # Note: In doc_to_webdyn.py, we apply address_offset during the final generation stage
    # to maintain consistency with the old behavior of passing raw mapped data.
    mapped = extractor.map_and_clean(raw, address_offset=0)
    if not mapped:
        logging.error("No registers extracted.")
        sys.exit(1)

    generator = Generator()
    processed = generator.process_rows(mapped, args.address_offset)

    if not args.output:
        mfg_clean = RE_TAG_CLEAN.sub('_', args.manufacturer).lower()
        model_clean = RE_TAG_CLEAN.sub('_', args.model).lower()
        output_file = f"{mfg_clean}_{model_clean}_definition.csv"
    else:
        output_file = args.output

    generator.write_output_csv(output_file, processed, args.manufacturer, args.model,
                               args.protocol, args.category, args.forced_write)

if __name__ == "__main__":
    main()
