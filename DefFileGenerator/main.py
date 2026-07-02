#!/usr/bin/env python3
import argparse
import sys
import os

# Ensure the parent directory is in sys.path to allow direct execution
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import logging
import csv
import json
import tempfile
from DefFileGenerator.extractor import Extractor, peek_generator
from DefFileGenerator.def_gen import Generator, run_generator, GeneratorConfig

def setup_logging(verbose=False):
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format='%(levelname)s: %(message)s',
        force=True
    )

def _perform_extraction(args):
    input_file = getattr(args, 'input_file', None)
    if not input_file:
        return None

    mapping = {}
    if getattr(args, 'mapping', None):
        try:
            with open(mapping_path, 'r') as f:
                mapping = json.load(f)
        except (OSError, ValueError) as e:
            logging.error(f"Error reading mapping file: {e}")
            sys.exit(1)

    extractor = Extractor(mapping)
    if not os.path.exists(input_file):
        logging.error(f"Input file not found: {input_file}")
        sys.exit(1)

    ext = os.path.splitext(input_file)[1].lower()
    address_offset = getattr(args, 'address_offset', 0)

    sheet = getattr(args, 'sheet', None)
    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw_data = extractor.extract_from_excel(input_file, getattr(args, 'sheet', None))
    elif ext == '.pdf':
        raw_data = extractor.extract_from_pdf(input_file, pages_arg)
    elif ext == '.csv':
        raw_data = extractor.extract_from_csv(input_file)
    elif ext == '.xml':
        raw_data = extractor.extract_from_xml(input_file)
    else:
        logging.error(f"Unsupported extension: {ext}")
        sys.exit(1)

    return extractor.map_and_clean(raw_data, address_offset)

def extract_command(args):
    mapped_data = _perform_extraction(args)
    is_empty, mapped_data = peek_generator(mapped_data)
    if is_empty:
        logging.error("No registers extracted.")
        sys.exit(1)

    output = getattr(args, 'output', None)
    fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']

    if output:
        f = open(output, 'w', newline='', encoding='utf-8')
    else:
        f = sys.stdout

    writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
    writer.writeheader()
    writer.writerows(mapped_data)

    if output:
        f.close()
        logging.info(f"Extraction complete. Saved to {output}")

def validate_command(args):
    generator = Generator()
    if generator.validate_csv(args.input_file):
        logging.info(f"Validation successful: {args.input_file}")
    else:
        logging.error(f"Validation failed: {args.input_file}")
        sys.exit(1)

def generate_command(args):
    template = getattr(args, 'template', False)
    template_mode = 'input'
    if template and args.input_file == 'definition':
        template_mode = 'definition'

    config = GeneratorConfig(
        input_file=getattr(args, 'input_file', None),
        output=args.output,
        manufacturer=getattr(args, 'manufacturer', 'Manufacturer'),
        model=getattr(args, 'model', 'Model'),
        protocol=getattr(args, 'protocol', 'modbusRTU'),
        category=getattr(args, 'category', 'Inverter'),
        forced_write=getattr(args, 'forced_write', ''),
        address_offset=getattr(args, 'address_offset', 0),
        template=template,
        template_mode=template_mode
    )
    run_generator(config)

def validate_command(args):
    generator = Generator()
    if not generator.validate_csv(args.input_file):
        logging.error(f"Validation failed for {args.input_file}")
        sys.exit(1)
    logging.info(f"Validation successful for {args.input_file}")

def run_command(args):
    template = getattr(args, 'template', False)
    if template:
        generate_command(args)
        return

    mapped_data = _perform_extraction(args)
    is_empty, mapped_data = peek_generator(mapped_data)
    if is_empty:
        logging.error("No registers extracted.")
        sys.exit(1)

    config = GeneratorConfig(
        input_file=getattr(args, 'input_file', None),
        output=args.output,
        manufacturer=getattr(args, 'manufacturer', 'Manufacturer'),
        model=getattr(args, 'model', 'Model'),
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        address_offset=0, # Already applied during extraction in run mode
        template=template_flag
    )
    run_generator(config, input_data=mapped_data if not template_flag else None)

def validate_command(args):
    generator = Generator()
    if not generator.validate_csv(args.input_file):
        sys.exit(1)

def validate_command(args):
    generator = Generator()
    if generator.validate_csv(args.input_file):
        logging.info(f"Validation successful: {args.input_file}")
    else:
        logging.error(f"Validation failed: {args.input_file}")
        sys.exit(1)

def validate_command(args):
    generator = Generator()
    if generator.validate_csv(args.input_file):
        logging.info(f"Validation successful for {args.input_file}")
    else:
        logging.error(f"Validation failed for {args.input_file}")
        sys.exit(1)

def validate_command(args):
    generator = Generator()
    if not generator.validate_csv(args.input_file):
        sys.exit(1)

def _run_cli():
    parser = argparse.ArgumentParser(description='WebdynSunPM Definition Tool')
    parser.add_argument('-v', '--verbose', action='store_true', help='Verbose logging')
    subparsers = parser.add_subparsers(dest='command', help='Sub-commands')

    # Validate
    parser_validate = subparsers.add_parser('validate', help='Validate a definition file')
    parser_validate.add_argument('input_file', help='Definition CSV to validate')

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
    parser_generate.add_argument('input_file', nargs='?', help='Input CSV')
    parser_generate.add_argument('--manufacturer')
    parser_generate.add_argument('--model')
    parser_generate.add_argument('-o', '--output', help='Output definition CSV')
    parser_generate.add_argument('--protocol', default='modbusRTU')
    parser_generate.add_argument('--category', default='Inverter')
    parser_generate.add_argument('--forced-write', default='')
    parser_generate.add_argument('--template', action='store_true', help='Generate a template CSV')
    parser_generate.add_argument('--address-offset', type=int, default=0, help='Address offset')
    parser_generate.add_argument('--template', action='store_true', help='Generate a template')

    # Run (Extract + Generate)
    parser_run = subparsers.add_parser('run', help='Extract and Generate in one step')
    parser_run.add_argument('input_file', nargs='?', help='Source file (PDF/Excel/CSV/XML)')
    parser_run.add_argument('--manufacturer')
    parser_run.add_argument('--model')
    parser_run.add_argument('-o', '--output', help='Output definition CSV')
    parser_run.add_argument('--mapping', help='Mapping JSON')
    parser_run.add_argument('--sheet', help='Excel sheet')
    parser_run.add_argument('--pages', help='PDF pages')
    parser_run.add_argument('--protocol', default='modbusRTU')
    parser_run.add_argument('--category', default='Inverter')
    parser_run.add_argument('--forced-write', default='')
    parser_run.add_argument('--template', action='store_true', help='Generate a template CSV')
    parser_run.add_argument('--address-offset', type=int, default=0, help='Address offset')
    parser_run.add_argument('--template', action='store_true', help='Generate a template')

    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        return

    setup_logging(args.verbose)

    if args.command == 'extract':
        extract_command(args)
    elif args.command == 'generate':
        generate_command(args)
    elif args.command == 'run':
        run_command(args)
    elif args.command == 'validate':
        if not Generator.validate_csv(args.input_file):
            sys.exit(1)

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

if __name__ == '__main__':
    main()
