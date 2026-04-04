#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import csv
import json
import tempfile
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator, run_generator, GeneratorConfig

def setup_logging(verbose=False):
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format='%(levelname)s: %(message)s',
        force=True
    )

def _perform_extraction(args):
    if not hasattr(args, 'input_file') or not args.input_file:
        logging.error("input_file is required for extraction.")
        return []

    mapping = {}
    if hasattr(args, 'mapping') and args.mapping:
        try:
            with open(args.mapping, 'r') as f:
                mapping = json.load(f)
        except (json.JSONDecodeError, FileNotFoundError) as e:
            logging.error(f"Error reading mapping file: {e}")

    extractor = Extractor(mapping)
    ext = os.path.splitext(args.input_file)[1].lower()

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw_data = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        try:
            pages = [int(p.strip()) for p in args.pages.split(',')] if args.pages else None
        except ValueError:
            logging.error(f"Invalid --pages argument: {args.pages}. Use comma-separated integers.")
            pages = None
        raw_data = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv':
        raw_data = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml':
        raw_data = extractor.extract_from_xml(args.input_file)
    else:
        logging.error(f"Unsupported extension: {ext}")
        return []

    address_offset = getattr(args, 'address_offset', 0)
    return extractor.map_and_clean(raw_data, address_offset)

def extract_command(args):
    mapped_data = _perform_extraction(args)
    if not mapped_data:
        return

    output = args.output if args.output else sys.stdout
    fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']

    if isinstance(output, str):
        f = open(output, 'w', newline='', encoding='utf-8')
    else:
        f = output

    writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
    writer.writeheader()
    writer.writerows(mapped_data)

    if isinstance(output, str):
        f.close()
        logging.info(f"Extraction complete. Saved to {args.output}")

def generate_command(args):
    config = GeneratorConfig(
        input_file=args.input_file,
        output=args.output,
        manufacturer=args.manufacturer,
        model=args.model,
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        template=args.template,
        address_offset=args.address_offset
    )
    run_generator(config)

def run_command(args):
    # For 'run', extraction handles the offset, so generator should use 0 to avoid double offset
    mapped_data = _perform_extraction(args)
    if not mapped_data:
        return

    with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8', newline='') as tf:
        temp_csv = tf.name
        fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
        writer = csv.DictWriter(tf, fieldnames=fieldnames, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(mapped_data)

    try:
        config = GeneratorConfig(
            input_file=temp_csv,
            output=args.output,
            manufacturer=args.manufacturer,
            model=args.model,
            protocol=args.protocol,
            category=args.category,
            forced_write=args.forced_write,
            address_offset=0 # Already applied in _perform_extraction
        )
        run_generator(config)
    finally:
        if os.path.exists(temp_csv):
            os.remove(temp_csv)

def main():
    # Pre-parse for verbose to set up logging early
    verbose = '-v' in sys.argv or '--verbose' in sys.argv
    setup_logging(verbose)
    parser = argparse.ArgumentParser(description='WebdynSunPM Definition Tool')
    subparsers = parser.add_subparsers(dest='command', help='Sub-commands')

    # Shared arguments
    global_parser = argparse.ArgumentParser(add_help=False)
    global_parser.add_argument('-v', '--verbose', action='store_true')
    global_parser.add_argument('--address-offset', type=int, default=0)

    # Extract
    parser_extract = subparsers.add_parser('extract', parents=[global_parser], help='Extract registers from documentation')
    parser_extract.add_argument('input_file', help='Source file (PDF/Excel/CSV/XML)')
    parser_extract.add_argument('-o', '--output', help='Output CSV')
    parser_extract.add_argument('--mapping', help='Mapping JSON')
    parser_extract.add_argument('--sheet', help='Excel sheet')
    parser_extract.add_argument('--pages', help='PDF pages')

    # Generate
    parser_generate = subparsers.add_parser('generate', parents=[global_parser], help='Generate definition from CSV')
    parser_generate.add_argument('input_file', nargs='?', help='Input CSV')
    parser_generate.add_argument('--manufacturer', help='Manufacturer name')
    parser_generate.add_argument('--model', help='Model name')
    parser_generate.add_argument('-o', '--output', help='Output definition CSV')
    parser_generate.add_argument('--protocol', default='modbusRTU')
    parser_generate.add_argument('--category', default='Inverter')
    parser_generate.add_argument('--forced-write', default='')
    parser_generate.add_argument('--template', action='store_true')

    # Run (Extract + Generate)
    parser_run = subparsers.add_parser('run', parents=[global_parser], help='Extract and Generate in one step')
    parser_run.add_argument('input_file', help='Source file (PDF/Excel/CSV/XML)')
    parser_run.add_argument('--manufacturer', required=True)
    parser_run.add_argument('--model', required=True)
    parser_run.add_argument('-o', '--output', help='Output definition CSV')
    parser_run.add_argument('--mapping', help='Mapping JSON')
    parser_run.add_argument('--sheet', help='Excel sheet')
    parser_run.add_argument('--pages', help='PDF pages')
    parser_run.add_argument('--protocol', default='modbusRTU')
    parser_run.add_argument('--category', default='Inverter')
    parser_run.add_argument('--forced-write', default='')

    args = parser.parse_args()

    # Manual validation for required manufacturer/model in non-template mode
    if args.command == 'generate' and not args.template:
        if not args.manufacturer or not args.model or not args.input_file:
            parser_generate.error("the following arguments are required: input_file, --manufacturer, --model")

    if args.command == 'extract':
        extract_command(args)
    elif args.command == 'generate':
        generate_command(args)
    elif args.command == 'run':
        run_command(args)
    else:
        parser.print_help()

if __name__ == '__main__':
    main()
