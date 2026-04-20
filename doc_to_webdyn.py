#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
import json
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator

def main():
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser: Extract registers from PDF, Excel, CSV, or XML and generate definition files.')
    parser.add_argument('input_file', help='Path to manufacturer documentation (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer', required=True, help='Manufacturer name (e.g., Huawei)')
    parser.add_argument('--model', required=True, help='Model name (e.g., SUN2000-5KTL)')
    parser.add_argument('-o', '--output', help='Output filename (default: mfg_model_definition.csv)')
    parser.add_argument('--protocol', default='modbusRTU', help='Protocol name (default: modbusRTU)')
    parser.add_argument('--category', default='Inverter', help='Device category (default: Inverter)')
    parser.add_argument('--sheet', help='Excel sheet name (processes all or active if omitted)')
    parser.add_argument('--mapping', help='Path to JSON file mapping manufacturer columns to internal ones (e.g., {"Address": "Modbus Addr"})')
    parser.add_argument('--pages', help='Comma-separated list of PDF pages to extract from (e.g., "1,2,5")')
    parser.add_argument('--address-offset', type=int, default=0, help='Integer value to add to all extracted register addresses')
    parser.add_argument('--forced-write', default='', help='Value for the Forced Write column in the definition header')
    parser.add_argument('-v', '--verbose', action='store_true', help='Enable detailed debug logging')

    args = parser.parse_args()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format='%(levelname)s: %(message)s', force=True)

    mapping = {}
    if args.mapping:
        try:
            with open(args.mapping, 'r') as f:
                mapping = json.load(f)
        except Exception as e:
            logging.error(f"Error loading mapping file: {e}")
            sys.exit(1)

    ext = os.path.splitext(args.input_file)[1].lower()
    extractor = Extractor(mapping)

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = None
        if args.pages:
            try:
                pages = [int(p.strip()) for p in args.pages.split(',') if p.strip()]
            except ValueError:
                logging.error(f"Invalid --pages format: {args.pages}. Use comma-separated integers (e.g., 1,2,5).")
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
        logging.error("No data extracted. Verify your source file and parameters.")
        sys.exit(1)

    mapped = extractor.map_and_clean(raw)
    if not mapped:
        logging.error("No registers extracted. Could not identify any register tables.")
        sys.exit(1)

    generator = Generator()
    processed = generator.process_rows(mapped, address_offset=args.address_offset)

    output_file = args.output or f"{re.sub(r'[^a-zA-Z0-9]', '_', args.manufacturer).lower()}_{re.sub(r'[^a-zA-Z0-9]', '_', args.model).lower()}_definition.csv"
    generator.write_output_csv(output_file, processed, args.manufacturer, args.model, args.protocol, args.category, args.forced_write)

if __name__ == "__main__":
    main()
