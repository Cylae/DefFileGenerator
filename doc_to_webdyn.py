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
    setup_logging(args.verbose)

    if args.template:
        config = GeneratorConfig(
            output=args.output,
            template=True,
            template_mode=args.template_mode
        )
        run_generator(config)
        return

    if not args.input_file or not os.path.exists(args.input_file):
        logging.error(f"Input file not found: {args.input_file}")
        sys.exit(1)

    mapping = {}
    if args.mapping:
        try:
            with open(args.mapping, 'r') as f:
                mapping = json.load(f)
        except (OSError, ValueError) as e:
            logging.error(f"Error reading mapping file: {e}")
            sys.exit(1)

    extractor = Extractor(mapping)
    ext = os.path.splitext(args.input_file)[1].lower()

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']: raw = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf': raw = extractor.extract_from_pdf(args.input_file, args.pages)
    elif ext == '.csv': raw = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml': raw = extractor.extract_from_xml(args.input_file)
    else: logging.error(f"Unsupported extension: {ext}"); sys.exit(1)

    has_data, raw_peeked = peek_generator(raw)
    if not has_data: logging.error("No data extracted."); sys.exit(1)

    mapped = extractor.map_and_clean(raw_peeked, args.address_offset)
    has_regs, mapped_peeked = peek_generator(mapped)
    if not has_regs: logging.error("No registers extracted."); sys.exit(1)

    m_name = args.manufacturer or "Manufacturer"
    m_model = args.model or "Model"
    output_file = args.output or f"{re.sub(r'[^a-zA-Z0-9]', '_', m_name).lower()}_{re.sub(r'[^a-zA-Z0-9]', '_', m_model).lower()}_definition.csv"

    config = GeneratorConfig(
        output=output_file,
        manufacturer=m_name,
        model=m_model,
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        address_offset=0 # Already applied during extraction
    )
    run_generator(config, input_data=mapped_peeked)

def setup_logging(verbose):
    logging.basicConfig(level=logging.DEBUG if verbose else logging.INFO, format='%(levelname)s: %(message)s', force=True)

def main():
    try:
        _run_cli()
    except KeyboardInterrupt:
        sys.exit(130)
    except SystemExit:
        raise
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
