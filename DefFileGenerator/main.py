#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import csv
import json

# Ensure the parent directory is in sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator, run_generator, GeneratorConfig, peek_generator

def setup_logging(verbose=False):
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format='%(levelname)s: %(message)s',
        force=True
    )

def _perform_extraction(args):
    input_file = getattr(args, 'input_file', None)
    if not input_file or not os.path.exists(input_file):
        logging.error(f"Input file not found: {input_file}")
        sys.exit(1)

    mapping = {}
    if getattr(args, 'mapping', None):
        try:
            with open(args.mapping, 'r') as f:
                mapping = json.load(f)
        except (OSError, ValueError) as e:
            logging.error(f"Error reading mapping file: {e}")
            sys.exit(1)

    extractor = Extractor(mapping)
    ext = os.path.splitext(input_file)[1].lower()

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw_data = extractor.extract_from_excel(input_file, getattr(args, 'sheet', None))
    elif ext == '.pdf':
        raw_data = extractor.extract_from_pdf(input_file, getattr(args, 'pages', None))
    elif ext == '.csv':
        raw_data = extractor.extract_from_csv(input_file)
    elif ext == '.xml':
        raw_data = extractor.extract_from_xml(input_file)
    else:
        logging.error(f"Unsupported extension: {ext}")
        sys.exit(1)

    has_data, raw_peeked = peek_generator(raw_data)
    if not has_data:
        logging.error("No data extracted.")
        sys.exit(1)

    mapped_gen = extractor.map_and_clean(raw_peeked, getattr(args, 'address_offset', 0))
    has_regs, mapped_peeked = peek_generator(mapped_gen)
    if not has_regs:
        logging.error("No registers extracted.")
        sys.exit(1)

    return mapped_peeked

def extract_command(args):
    mapped_iterator = _perform_extraction(args)
    output = getattr(args, 'output', None)
    f = open(output, 'w', newline='', encoding='utf-8') if output else sys.stdout
    fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
    writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
    writer.writeheader()
    writer.writerows(mapped_iterator)
    if output:
        f.close()
        logging.info(f"Extraction complete. Saved to {output}")

def validate_command(args):
    if Generator().validate_csv(args.input_file):
        logging.info(f"Validation successful for {args.input_file}")
    else:
        logging.error(f"Validation failed for {args.input_file}")
        sys.exit(1)

def generate_command(args):
    config = GeneratorConfig(
        input_file=args.input_file,
        output=args.output,
        manufacturer=args.manufacturer,
        model=args.model,
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        address_offset=args.address_offset,
        template=args.template,
        template_mode=getattr(args, 'template_mode', 'input')
    )
    run_generator(config)

def run_command(args):
    template = getattr(args, 'template', False)
    mapped_data = None
    if not template:
        mapped_data = _perform_extraction(args)

    config = GeneratorConfig(
        input_file=args.input_file,
        output=args.output,
        manufacturer=args.manufacturer,
        model=args.model,
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        address_offset=0, # Already applied during extraction
        template=template
    )
    run_generator(config, input_data=mapped_data)

def _run_cli():
    parser = argparse.ArgumentParser(description='WebdynSunPM Definition Tool')
    parser.add_argument('-v', '--verbose', action='store_true', help='Verbose logging')
    subparsers = parser.add_subparsers(dest='command', help='Sub-commands')

    # Validate
    p_val = subparsers.add_parser('validate', help='Validate a definition file')
    p_val.add_argument('input_file', help='Definition CSV to validate')

    # Extract
    p_ext = subparsers.add_parser('extract', help='Extract registers from documentation')
    p_ext.add_argument('input_file', help='Source file (PDF/Excel/CSV/XML)')
    p_ext.add_argument('-o', '--output', help='Output CSV')
    p_ext.add_argument('--mapping', help='Mapping JSON')
    p_ext.add_argument('--sheet', help='Excel sheet')
    p_ext.add_argument('--pages', help='PDF pages')
    p_ext.add_argument('--address-offset', type=int, default=0)

    # Generate
    p_gen = subparsers.add_parser('generate', help='Generate definition from CSV')
    p_gen.add_argument('input_file', nargs='?', help='Input CSV')
    p_gen.add_argument('--manufacturer', default='Manufacturer')
    p_gen.add_argument('--model', default='Model')
    p_gen.add_argument('-o', '--output', help='Output definition CSV')
    p_gen.add_argument('--template', action='store_true')
    p_gen.add_argument('--template-mode', choices=['input', 'definition'], default='input')
    p_gen.add_argument('--protocol', default='modbusRTU')
    p_gen.add_argument('--category', default='Inverter')
    p_gen.add_argument('--forced-write', default='')
    p_gen.add_argument('--address-offset', type=int, default=0)

    # Run
    p_run = subparsers.add_parser('run', help='Extract and Generate in one step')
    p_run.add_argument('input_file', nargs='?', help='Source file')
    p_run.add_argument('--manufacturer', default='Manufacturer')
    p_run.add_argument('--model', default='Model')
    p_run.add_argument('-o', '--output', help='Output definition CSV')
    p_run.add_argument('--template', action='store_true')
    p_run.add_argument('--mapping', help='Mapping JSON')
    p_run.add_argument('--sheet', help='Excel sheet')
    p_run.add_argument('--pages', help='PDF pages')
    p_run.add_argument('--protocol', default='modbusRTU')
    p_run.add_argument('--category', default='Inverter')
    p_run.add_argument('--forced-write', default='')
    p_run.add_argument('--address-offset', type=int, default=0)

    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        return

    setup_logging(args.verbose)

    if args.command in ['generate', 'run'] and not getattr(args, 'template', False):
        if not getattr(args, 'input_file', None):
            logging.error("input_file is required.")
            sys.exit(1)

    if args.command == 'extract': extract_command(args)
    elif args.command == 'validate': validate_command(args)
    elif args.command == 'generate': generate_command(args)
    elif args.command == 'run': run_command(args)

def main():
    try:
        _run_cli()
    except KeyboardInterrupt:
        sys.exit(130)
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()
