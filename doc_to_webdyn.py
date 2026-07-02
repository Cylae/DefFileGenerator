#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
import csv
from DefFileGenerator.extractor import Extractor, peek_generator
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator

def _run_cli():
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', help='Path to documentation (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer', default='Manufacturer')
    parser.add_argument('--model', default='Model')
    parser.add_argument('-o', '--output', help='Output filename')
    parser.add_argument('--protocol', default='modbusRTU')
    parser.add_argument('--category', default='Inverter')
    parser.add_argument('--sheet', help='Excel sheet name')
    parser.add_argument('--pages', help='PDF pages (comma-separated integers)')
    parser.add_argument('--mapping', help='JSON mapping file')
    parser.add_argument('--address-offset', type=int, default=0)
    parser.add_argument('--forced-write', default='')
    parser.add_argument('--template', action='store_true', help='Generate template CSV')
    parser.add_argument('-v', '--verbose', action='store_true')

    args = parser.parse_args()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format='%(levelname)s: %(message)s', force=True)

    if args.template:
        config = GeneratorConfig(
            output=args.output,
            manufacturer=args.manufacturer,
            model=args.model,
            template=True
        )
        run_generator(config)
        return

    if not args.input_file:
        logging.error("input_file is required when not using --template.")
        sys.exit(1)

    if not args.manufacturer or not args.model:
        logging.error("--manufacturer and --model are required.")
        sys.exit(1)

    if not os.path.exists(args.input_file):
        logging.error(f"Input file not found: {args.input_file}")
        sys.exit(1)

    ext = os.path.splitext(args.input_file)[1].lower()

    mapping = {}
    mapping_arg = getattr(args, 'mapping', None)
    if mapping_arg:
        try:
            with open(mapping_arg, 'r') as f:
                mapping = json.load(f)
        except (OSError, ValueError) as e:
            logging.error(f"Error reading mapping file: {e}")
            sys.exit(1)

    extractor = Extractor(mapping)

    pages = None
    pages_arg = getattr(args, 'pages', None)
    if pages_arg:
        if ext != '.pdf':
            logging.warning("--pages is only applicable for PDF files. Ignoring.")
        else:
            try:
                pages = [int(p.strip()) for p in pages_arg.split(',')]
            except ValueError:
                logging.error("Invalid format for --pages. Expected comma-separated integers.")
                sys.exit(1)

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']: raw = extractor.extract_from_excel(args.input_file, getattr(args, 'sheet', None))
    elif ext == '.pdf': raw = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv': raw = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml': raw = extractor.extract_from_xml(args.input_file)
    else: logging.error(f"Unsupported extension: {ext}"); sys.exit(1)

    is_empty, raw = peek_generator(raw)
    if is_empty: logging.error("No data extracted."); sys.exit(1)

    mapped = extractor.map_and_clean(raw, getattr(args, 'address_offset', 0))
    exists, mapped = peek_generator(mapped)
    if not exists: logging.error("No registers extracted."); sys.exit(1)

    manufacturer = getattr(args, 'manufacturer', 'Manufacturer')
    model = getattr(args, 'model', 'Model')
    output_file = getattr(args, 'output', None) or f"{re.sub(r'[^a-zA-Z0-9]', '_', manufacturer).lower()}_{re.sub(r'[^a-zA-Z0-9]', '_', model).lower()}_definition.csv"

    config = GeneratorConfig(
        input_file=args.input_file,
        output=output_file,
        manufacturer=manufacturer,
        model=model,
        protocol=getattr(args, 'protocol', 'modbusRTU'),
        category=getattr(args, 'category', 'Inverter'),
        forced_write=getattr(args, 'forced_write', ''),
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
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
