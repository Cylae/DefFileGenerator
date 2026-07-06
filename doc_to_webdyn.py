#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
import csv
import json
from DefFileGenerator.extractor import Extractor, peek_generator
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator

def _run_cli():
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', nargs='?', help='Path to documentation (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer', help='Manufacturer name')
    parser.add_argument('--model', help='Model name')
    parser.add_argument('--template', action='store_true', help='Generate a template definition')
    parser.add_argument('--template-mode', choices=['input', 'definition'], default='input')
    parser.add_argument('-o', '--output', help='Output filename')
    parser.add_argument('--protocol', default='modbusRTU')
    parser.add_argument('--category', default='Inverter')
    parser.add_argument('--sheet', help='Excel sheet name')
    parser.add_argument('--pages', help='PDF pages (comma-separated integers)')
    parser.add_argument('--mapping', help='JSON mapping file')
    parser.add_argument('--address-offset', type=int, default=0)
    parser.add_argument('--forced-write', default='')
    parser.add_argument('-v', '--verbose', action='store_true')

    args = parser.parse_args()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format='%(levelname)s: %(message)s', force=True)

    if args.template:
        config = GeneratorConfig(output=args.output, template=True, template_mode=args.template_mode)
        run_generator(config)
        return

    if not args.input_file or not os.path.exists(args.input_file):
        logging.error(f"Input file not found: {args.input_file}")
        sys.exit(1)

    ext = os.path.splitext(args.input_file)[1].lower()

    mapping = {}
    if args.mapping:
        if not os.path.exists(args.mapping):
            logging.error(f"Mapping file not found: {args.mapping}")
            sys.exit(1)
        try:
            with open(args.mapping, 'r') as f:
                mapping = json.load(f)
        except Exception as e:
            logging.error(f"Error reading mapping file: {e}")
            sys.exit(1)

    extractor = Extractor(mapping)

    if ext in ['.xlsx', '.xlsm']: raw = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf': raw = extractor.extract_from_pdf(args.input_file, args.pages)
    elif ext == '.csv': raw = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml': raw = extractor.extract_from_xml(args.input_file)
    else:
        logging.error(f"Unsupported extension: {ext}")
        sys.exit(1)

    has_data, raw_peeked = peek_generator(raw)
    if not has_data:
        logging.error("No data extracted.")
        sys.exit(1)

    mapped = extractor.map_and_clean(raw_peeked, args.address_offset)
    has_regs, mapped_peeked = peek_generator(mapped)
    if not has_regs:
        logging.error("No registers extracted.")
        sys.exit(1)

    m_name = args.manufacturer or "Manufacturer"
    m_model = args.model or "Model"
    output_file = args.output or f"{re.sub(r'[^a-zA-Z0-9]', '_', m_name).lower()}_{re.sub(r'[^a-zA-Z0-9]', '_', m_model).lower()}_definition.csv"

    config = GeneratorConfig(
        input_file=args.input_file,
        output=output_file,
        manufacturer=m_name,
        model=m_model,
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        address_offset=0 # Offset already applied in map_and_clean
    )
    run_generator(config, input_data=mapped_peeked)

def main():
    try:
        _run_cli()
    except KeyboardInterrupt:
        sys.exit(130)
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
