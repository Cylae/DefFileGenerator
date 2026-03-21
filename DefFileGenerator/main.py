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

def setup_logging():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def _perform_extraction(args, address_offset=0):
    mapping = {}
    if hasattr(args, 'mapping') and args.mapping:
        try:
            with open(args.mapping, 'r') as f:
                mapping = json.load(f)
        except (Exception, json.JSONDecodeError) as e:
            logging.error(f"Error reading mapping file: {e}")

    extractor = Extractor(mapping)
    ext = os.path.splitext(args.input_file)[1].lower()

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw_data = extractor.extract_from_excel(args.input_file, getattr(args, 'sheet', None))
    elif ext == '.pdf':
        pages = None
        if hasattr(args, 'pages') and args.pages:
            try:
                pages = [int(p.strip()) for p in args.pages.split(',')]
            except ValueError:
                logging.error(f"Invalid pages format: {args.pages}. Use comma-separated integers.")
        raw_data = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv':
        raw_data = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml':
        raw_data = extractor.extract_from_xml(args.input_file)
    else:
        logging.error(f"Unsupported extension: {ext}")
        return []

    return extractor.map_and_clean(raw_data, address_offset=address_offset, args=args)

def extract_command(args):
    mapped_data = _perform_extraction(args, address_offset=args.address_offset)
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
        address_offset=args.address_offset
    )
    run_generator(config)

def run_command(args):
    # Apply offset during extraction
    mapped_data = _perform_extraction(args, address_offset=args.address_offset)
    if not mapped_data:
        return

    with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as tf:
        temp_csv = tf.name
        fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
        writer = csv.DictWriter(tf, fieldnames=fieldnames, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(mapped_data)

    try:
        # Pass 0 to generator to avoid double-offset
        config = GeneratorConfig(
            input_file=temp_csv,
            output=args.output,
            manufacturer=args.manufacturer,
            model=args.model,
            protocol=args.protocol,
            category=args.category,
            forced_write=args.forced_write,
            address_offset=0
        )
        run_generator(config)
    finally:
        if os.path.exists(temp_csv):
            os.remove(temp_csv)

def main():
    setup_logging()
    parser = argparse.ArgumentParser(description='WebdynSunPM Definition Tool')
    subparsers = parser.add_subparsers(dest='command', help='Sub-commands')

    # Extract
    parser_extract = subparsers.add_parser('extract', help='Extract registers from documentation')
    parser_extract.add_argument('input_file', help='Source file (PDF/Excel/CSV/XML)')
    parser_extract.add_argument('-o', '--output', help='Output CSV')
    parser_extract.add_argument('--mapping', help='Mapping JSON')
    parser_extract.add_argument('--sheet', help='Excel sheet')
    parser_extract.add_argument('--pages', help='PDF pages')
    parser_extract.add_argument('--address-offset', type=int, default=0, help='Shift addresses by offset')

    # Generate
    parser_generate = subparsers.add_parser('generate', help='Generate definition from CSV')
    parser_generate.add_argument('input_file', help='Input CSV')
    parser_generate.add_argument('--manufacturer', required=True)
    parser_generate.add_argument('--model', required=True)
    parser_generate.add_argument('-o', '--output', help='Output definition CSV')
    parser_generate.add_argument('--protocol', default='modbusRTU')
    parser_generate.add_argument('--category', default='Inverter')
    parser_generate.add_argument('--forced-write', default='')
    parser_generate.add_argument('--address-offset', type=int, default=0, help='Shift addresses by offset')

    # Run (Extract + Generate)
    parser_run = subparsers.add_parser('run', help='Extract and Generate in one step')
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
    parser_run.add_argument('--address-offset', type=int, default=0, help='Shift addresses by offset')

    args = parser.parse_args()

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
