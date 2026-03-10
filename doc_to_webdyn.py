#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator

def main():
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', help='Path to documentation (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer', required=True)
    parser.add_argument('--model', required=True)
    parser.add_argument('-o', '--output', help='Output filename')
    parser.add_argument('--protocol', default='modbusRTU')
    parser.add_argument('--category', default='Inverter')
    parser.add_argument('--sheet', help='Excel sheet name')
    parser.add_argument('--pages', help='PDF pages (comma-separated)')
    parser.add_argument('--mapping', help='Mapping JSON file')
    parser.add_argument('--address-offset', type=int, default=0)
    parser.add_argument('--forced-write', default='')
    parser.add_argument('-v', '--verbose', action='store_true')

    args = parser.parse_args()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format='%(levelname)s: %(message)s', force=True)

    from DefFileGenerator.main import _perform_extraction
    mapped = _perform_extraction(args)

    if not mapped:
        logging.error("No registers extracted.")
        sys.exit(1)

    generator = Generator()
    processed = generator.process_rows(mapped, args.address_offset)

    output_file = args.output or f"{re.sub(r'[^a-zA-Z0-9]', '_', args.manufacturer).lower()}_{re.sub(r'[^a-zA-Z0-9]', '_', args.model).lower()}_definition.csv"
    generator.write_output_csv(output_file, processed, args.manufacturer, args.model, args.protocol, args.category, args.forced_write)

if __name__ == "__main__":
    main()
