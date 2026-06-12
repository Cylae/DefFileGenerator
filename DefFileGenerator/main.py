#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import csv
import json
import re
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator, run_generator, GeneratorConfig, peek_generator

def setup_logging(verbose=False):
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format='%(levelname)s: %(message)s',
        force=True
    )

def _perform_extraction(args):
    mapping = {}
    mapping_file = getattr(args, 'mapping', None)
    if mapping_file:
        try:
            with open(mapping_file, 'r') as f:
                mapping = json.load(f)
        except (OSError, ValueError) as e:
            logging.error(f"Error reading mapping file: {e}")
            sys.exit(1)

    extractor = Extractor(mapping)
    input_file = getattr(args, 'input_file', None)
    if not input_file or not os.path.exists(input_file):
        logging.error(f"Input file not found: {input_file}")
        sys.exit(1)

    ext = os.path.splitext(input_file)[1].lower()
    address_offset = getattr(args, 'address_offset', 0)
    pages_arg = getattr(args, 'pages', None)
    sheet_arg = getattr(args, 'sheet', None)

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw_data = extractor.extract_from_excel(input_file, sheet_arg)
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
    mapped_it = _perform_extraction(args)
    has_data, restored_it = peek_generator(mapped_it)
    if not has_data:
        logging.error("No registers extracted.")
        sys.exit(1)

    output = args.output if args.output else sys.stdout
    fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']

    if isinstance(output, str):
        f = open(output, 'w', newline='', encoding='utf-8')
    else:
        f = output

    writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
    writer.writeheader()
    writer.writerows(restored_it)

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
        address_offset=args.address_offset,
        template=args.template
    )
    run_generator(config)

def validate_command(args):
    generator = Generator()
    if generator.validate_csv(args.input_file):
        logging.info(f"Validation successful: {args.input_file}")
    else:
        logging.error(f"Validation failed: {args.input_file}")
        sys.exit(1)

def run_command(args):
    if args.template:
        config = GeneratorConfig(template=True, output=args.output)
        run_generator(config)
        return

    mapped_it = _perform_extraction(args)
    has_data, restored_it = peek_generator(mapped_it)
    if not has_data:
        logging.error("No registers extracted.")
        sys.exit(1)

    config = GeneratorConfig(
        input_file=args.input_file,
        output=args.output,
        manufacturer=args.manufacturer,
        model=args.model,
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        address_offset=0 # Already applied during extraction in run mode
    )
    run_generator(config, input_data=restored_it)

def _run_cli():
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
    parser_generate.add_argument('input_file', nargs='?', help='Input CSV')
    parser_generate.add_argument('--manufacturer')
    parser_generate.add_argument('--model')
    parser_generate.add_argument('-o', '--output', help='Output definition CSV')
    parser_generate.add_argument('--protocol', default='modbusRTU')
    parser_generate.add_argument('--category', default='Inverter')
    parser_generate.add_argument('--forced-write', default='')
    parser_generate.add_argument('--address-offset', type=int, default=0, help='Address offset')
    parser_generate.add_argument('--template', action='store_true', help='Generate template')

    # Validate
    parser_validate = subparsers.add_parser('validate', help='Validate a definition file')
    parser_validate.add_argument('input_file', help='Webdyn definition CSV to validate')

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
    parser_run.add_argument('--address-offset', type=int, default=0, help='Address offset')
    parser_run.add_argument('--template', action='store_true', help='Generate template')

    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        return

    setup_logging(args.verbose)

    # Validate --manufacturer and --model for generate/run unless template
    if args.command in ['generate', 'run'] and not args.template:
        if not args.input_file:
            parser.error(f"input_file is required for '{args.command}' unless --template is used.")
        if not args.manufacturer or not args.model:
            logging.error("Manufacturer and Model are required.")
            sys.exit(1)

    if args.command == 'extract':
        extract_command(args)
    elif args.command == 'generate':
        generate_command(args)
    elif args.command == 'validate':
        validate_command(args)
    elif args.command == 'run':
        run_command(args)

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

if __name__ == '__main__':
    main()
