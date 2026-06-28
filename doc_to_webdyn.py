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

    if getattr(args, 'template', False):
        config = GeneratorConfig(output=args.output, template=True)
        run_generator(config)
        return

    if not args.input_file or not os.path.exists(args.input_file):
        logging.error(f"Input file not found: {args.input_file}")
        sys.exit(1)

    ext = os.path.splitext(args.input_file)[1].lower()

    mapping = {}
    if args.mapping:
        try:
            with open(args.mapping, 'r') as f:
                mapping = json.load(f)
        except (OSError, ValueError) as e:
            logging.error(f"Error reading mapping file: {e}")
            sys.exit(1)

    extractor = Extractor(mapping)

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']: raw = extractor.extract_from_excel(args.input_file, getattr(args, 'sheet', None))
    elif ext == '.pdf': raw = extractor.extract_from_pdf(args.input_file, getattr(args, 'pages', None))
    elif ext == '.csv': raw = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml': raw = extractor.extract_from_xml(args.input_file)
    else: logging.error(f"Unsupported extension: {ext}"); sys.exit(1)

    is_not_empty, raw = peek_generator(raw)
    if not is_not_empty: logging.error("No data extracted."); sys.exit(1)

    mapped = extractor.map_and_clean(raw, args.address_offset)
    is_not_empty, mapped = peek_generator(mapped)
    if not is_not_empty: logging.error("No registers extracted."); sys.exit(1)

    manufacturer = getattr(args, 'manufacturer', 'Unknown')
    model = getattr(args, 'model', 'Model')
    output_file = args.output or f"{re.sub(r'[^a-zA-Z0-9]', '_', manufacturer).lower()}_{re.sub(r'[^a-zA-Z0-9]', '_', model).lower()}_definition.csv"

    config = GeneratorConfig(
        input_file=args.input_file,
        output=output_file,
        manufacturer=manufacturer,
        model=model,
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        address_offset=0 # Already applied during extraction
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
