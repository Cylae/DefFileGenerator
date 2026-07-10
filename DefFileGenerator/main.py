#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import csv
import json

# Ensure the parent directory is in sys.path to allow direct execution
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
    if not input_file:
        logging.error("Input file is required for extraction.")
        sys.exit(1)
    if not os.path.exists(input_file):
        logging.error(f"Input file not found: {input_file}")
        sys.exit(1)

    mapping = {}
    mapping_path = getattr(args, 'mapping', None)
    if mapping_path:
        try:
            with open(mapping_path, 'r') as f:
                mapping = json.load(f)
        except (OSError, ValueError) as e:
            logging.error(f"Error reading mapping file: {e}")
            sys.exit(1)

    extractor = Extractor(mapping)
    ext = os.path.splitext(input_file)[1].lower()
    address_offset = getattr(args, 'address_offset', 0)
    pages = getattr(args, 'pages', None)
    sheet = getattr(args, 'sheet', None)

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw_data = extractor.extract_from_excel(input_file, sheet)
    elif ext == '.pdf':
        raw_data = extractor.extract_from_pdf(input_file, pages)
    elif ext == '.csv':
        raw_data = extractor.extract_from_csv(input_file)
    elif ext == '.xml':
        raw_data = extractor.extract_from_xml(input_file)
    else:
        logging.error(f"Unsupported extension: {ext}")
        sys.exit(1)

    has_data, raw_data_peeked = peek_generator(raw_data)
    if not has_data:
        logging.error("No data extracted.")
        sys.exit(1)

    mapped_gen = extractor.map_and_clean(raw_data_peeked, address_offset)
    has_regs, mapped_peeked = peek_generator(mapped_gen)
    if not has_regs:
        logging.error("No registers extracted.")
        sys.exit(1)

    return list(mapped_peeked)

def extract_command(args):
    mapped_data = _perform_extraction(args)
    output = getattr(args, 'output', None)
    f = open(output, 'w', newline='', encoding='utf-8') if output else sys.stdout
    fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
    writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
    writer.writeheader()
    writer.writerows(mapped_data)
    if output:
        f.close()
        logging.info(f"Extraction complete. Saved to {output}")

def validate_command(args):
    generator = Generator()
    if generator.validate_csv(args.input_file):
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
        template_mode=args.template_mode
    )
    run_generator(config)

def run_command(args):
    mapped_data = None
    if not args.template:
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
        template=args.template,
        template_mode=args.template_mode
    )
    run_generator(config, input_data=mapped_data)

def _run_cli(args_list=None):
    parser = argparse.ArgumentParser(description='WebdynSunPM Definition Tool')
    parser.add_argument('-v', '--verbose', action='store_true', help='Verbose logging')
    subparsers = parser.add_subparsers(dest='command', help='Sub-commands')

    # Shared arguments
    def add_extraction_args(p):
        p.add_argument('--mapping', help='Mapping JSON')
        p.add_argument('--sheet', help='Excel sheet')
        p.add_argument('--pages', help='PDF pages')
        p.add_argument('--address-offset', type=int, default=0, help='Address offset')

    def add_generation_args(p):
        p.add_argument('--manufacturer')
        p.add_argument('--model')
        p.add_argument('--protocol', default='modbusRTU')
        p.add_argument('--category', default='Inverter')
        p.add_argument('--forced-write', default='')
        p.add_argument('--template', action='store_true')
        p.add_argument('--template-mode', choices=['input', 'definition'], default='input')

    # Validate
    parser_validate = subparsers.add_parser('validate', help='Validate a definition file')
    parser_validate.add_argument('input_file', help='Definition CSV to validate')

    # Extract
    parser_extract = subparsers.add_parser('extract', help='Extract registers from documentation')
    parser_extract.add_argument('input_file', help='Source file (PDF/Excel/CSV/XML)')
    parser_extract.add_argument('-o', '--output', help='Output CSV')
    add_extraction_args(parser_extract)

    # Generate
    parser_generate = subparsers.add_parser('generate', help='Generate definition from CSV')
    parser_generate.add_argument('input_file', nargs='?', help='Input CSV')
    parser_generate.add_argument('-o', '--output', help='Output definition CSV')
    add_generation_args(parser_generate)
    parser_generate.add_argument('--address-offset', type=int, default=0)

    # Run (Extract + Generate)
    parser_run = subparsers.add_parser('run', help='Extract and Generate in one step')
    parser_run.add_argument('input_file', nargs='?', help='Source file (PDF/Excel/CSV/XML)')
    parser_run.add_argument('-o', '--output', help='Output definition CSV')
    add_extraction_args(parser_run)
    add_generation_args(parser_run)

    args = parser.parse_args(args_list)

    if not args.command:
        parser.print_help()
        return

    setup_logging(args.verbose)

    # Required args validation
    if args.command in ['generate', 'run'] and not args.template:
        if not args.manufacturer or not args.model:
            logging.error("--manufacturer and --model are required.")
            sys.exit(1)
        if not args.input_file:
            logging.error("input_file is required.")
            sys.exit(1)

    if args.command == 'extract':
        extract_command(args)
    elif args.command == 'validate':
        validate_command(args)
    elif args.command == 'generate':
        generate_command(args)
    elif args.command == 'run':
        run_command(args)

def main(args=None):
    try:
        _run_cli(args)
    except KeyboardInterrupt:
        sys.exit(130)
    except SystemExit:
        raise
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()
