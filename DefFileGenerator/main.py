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
    mapping = {}
    if args.mapping:
        with open(args.mapping, 'r') as f:
            mapping = json.load(f)

    extractor = Extractor(mapping)
    ext = os.path.splitext(args.input_file)[1].lower()

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw_data = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = [int(p.strip()) for p in args.pages.split(',')] if args.pages else None
        raw_data = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv':
        raw_data = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml':
        raw_data = extractor.extract_from_xml(args.input_file)
    else:
        logging.error(f"Unsupported extension: {ext}")
        return []

    return extractor.map_and_clean(raw_data)

def extract_command(args):
    mapped_data = _perform_extraction(args)
    if not mapped_data:
        return

    # Apply address offset if specified
    if hasattr(args, 'address_offset') and args.address_offset != 0:
        generator = Generator()
        for row in mapped_data:
            row['Address'] = generator.apply_address_offset(row.get('Address', ''), args.address_offset)

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
    mapped_data = _perform_extraction(args)
    if not mapped_data:
        return

    with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as tf:
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
            address_offset=args.address_offset
        )
        run_generator(config)
    finally:
        if os.path.exists(temp_csv):
            os.remove(temp_csv)

def main():
    parser = argparse.ArgumentParser(description='WebdynSunPM Definition Tool')
    parser.add_argument('-v', '--verbose', action='store_true', help='Verbose logging')
    subparsers = parser.add_subparsers(dest='command', help='Sub-commands')

    # Extract
    parser_extract = subparsers.add_parser('extract', help='Extract registers from documentation')
    parser_extract.add_argument('input_file', help='Source file (PDF/Excel/CSV/XML)')
    parser_extract.add_argument('-o', '--output', help='Output CSV')
    parser_extract.add_argument('--mapping', help='Mapping JSON')
    parser_extract.add_argument('--sheet', help='Excel sheet')
    parser_extract.add_argument('--pages', help='PDF pages')
    parser_extract.add_argument('--address-offset', type=int, default=0, help='Address offset')

    # Generate
    parser_generate = subparsers.add_parser('generate', help='Generate definition from CSV')
    parser_generate.add_argument('input_file', help='Input CSV')
    parser_generate.add_argument('--manufacturer', required=True)
    parser_generate.add_argument('--model', required=True)
    parser_generate.add_argument('-o', '--output', help='Output definition CSV')
    parser_generate.add_argument('--protocol', default='modbusRTU')
    parser_generate.add_argument('--category', default='Inverter')
    parser_generate.add_argument('--forced-write', default='')
    parser_generate.add_argument('--address-offset', type=int, default=0, help='Address offset')

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
    parser_run.add_argument('--address-offset', type=int, default=0, help='Address offset')

    args = parser.parse_args()
    setup_logging(args.verbose)

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
