#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
import json
import csv
from DefFileGenerator.extractor import Extractor, peek_generator
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator

def _run_cli():
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', nargs='?', help='Path to documentation (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer')
    parser.add_argument('--model')
    parser.add_argument('-o', '--output', help='Output filename')
    parser.add_argument('--protocol', default='modbusRTU')
    parser.add_argument('--category', default='Inverter')
    parser.add_argument('--sheet', help='Excel sheet name')
    parser.add_argument('--pages', help='PDF pages (comma-separated integers)')
    parser.add_argument('--mapping', help='JSON mapping file')
    parser.add_argument('--address-offset', type=int, default=0)
    parser.add_argument('--forced-write', default='')
    parser.add_argument('--template', action='store_true')
    parser.add_argument('-v', '--verbose', action='store_true')

    args = parser.parse_args()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format='%(levelname)s: %(message)s', force=True)

    template_mode = getattr(args, 'template', False)

    if not template_mode:
        if not args.input_file:
            logging.error("input_file is required unless --template is used.")
            sys.exit(1)
        if not os.path.exists(args.input_file):
            logging.error(f"Input file not found: {args.input_file}")
            sys.exit(1)

    ext = os.path.splitext(args.input_file)[1].lower() if args.input_file else ""

    mapping = {}
    if args.mapping:
        try:
            with open(args.mapping, 'r') as f:
                mapping = json.load(f)
        except (OSError, ValueError) as e:
            logging.error(f"Error reading mapping file: {e}")
            sys.exit(1)

    extractor = Extractor(mapping)

    pages = None
    if args.pages:
        if ext != '.pdf':
            logging.warning("--pages is only applicable for PDF files. Ignoring.")
        else:
            try:
                pages = [int(p.strip()) for p in args.pages.split(',')]
            except ValueError:
                logging.error("Invalid format for --pages. Expected comma-separated integers.")
                sys.exit(1)

    mapped = None
    if not template_mode:
        if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']: raw = extractor.extract_from_excel(args.input_file, getattr(args, 'sheet', None))
        elif ext == '.pdf': raw = extractor.extract_from_pdf(args.input_file, pages)
        elif ext == '.csv': raw = extractor.extract_from_csv(args.input_file)
        elif ext == '.xml': raw = extractor.extract_from_xml(args.input_file)
        else: logging.error(f"Unsupported extension: {ext}"); sys.exit(1)

        has_data, raw_peeked = peek_generator(raw)
        if not has_data: logging.error("No data extracted."); sys.exit(1)

        mapped = list(extractor.map_and_clean(raw_peeked, args.address_offset))
        if not mapped: logging.error("No registers extracted."); sys.exit(1)

    mfg = getattr(args, 'manufacturer', 'Manufacturer')
    model = getattr(args, 'model', 'Model')
    output_file = args.output or f"{re.sub(r'[^a-zA-Z0-9]', '_', mfg).lower()}_{re.sub(r'[^a-zA-Z0-9]', '_', model).lower()}_definition.csv"

    config = GeneratorConfig(
        input_file=args.input_file,
        output=output_file,
        manufacturer=mfg,
        model=model,
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        address_offset=0 if not template_mode else getattr(args, 'address_offset', 0),
        template=template_mode
    )
    run_generator(config, input_data=mapped)

def main():
    try:
        _run_cli()
    except KeyboardInterrupt:
        sys.exit(130)
    except SystemExit:
        raise
    except (OSError, ValueError, TypeError, KeyError, csv.Error) as e:
        logging.error(f"An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
