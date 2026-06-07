#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import csv
import json
import tempfile
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator, run_generator, GeneratorConfig, peek_generator

def setup_logging(verbose=False):
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format='%(levelname)s: %(message)s'
    )

def _perform_extraction(args):
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
    if not os.path.exists(args.input_file):
        logging.error(f"Input file not found: {args.input_file}")
        sys.exit(1)

    ext = os.path.splitext(args.input_file)[1].lower()
    address_offset = getattr(args, 'address_offset', 0)
    pages_arg = getattr(args, 'pages', None)
    sheet_arg = getattr(args, 'sheet', None)

    if pages_arg and ext != '.pdf':
        logging.warning("--pages is only applicable for PDF files. Ignoring.")

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw_data = extractor.extract_from_excel(args.input_file, sheet_arg)
    elif ext == '.pdf':
        pages = None
        if pages_arg:
            try:
                pages = [int(p.strip()) for p in pages_arg.split(',')]
            except ValueError:
                logging.error("Invalid format for --pages. Expected comma-separated integers.")
                sys.exit(1)
        raw_data = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv':
        raw_data = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml':
        raw_data = extractor.extract_from_xml(args.input_file)
    else:
        logging.error(f"Unsupported extension: {ext}")
        sys.exit(1)

    return extractor.map_and_clean(raw_data, address_offset)

def extract_command(args):
    mapped_data = _perform_extraction(args)
    has_registers, mapped_data = peek_generator(mapped_data)
    if not has_registers:
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

def validate_command(args):
    if Generator.validate_csv(args.input_file):
        logging.info(f"Validation successful for {args.input_file}")
    else:
        logging.error(f"Validation failed for {args.input_file}")
        sys.exit(1)

def run_command(args):
    mapped_data = _perform_extraction(args)
    has_registers, mapped_data = peek_generator(mapped_data)
    if not has_registers:
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
    run_generator(config, input_data=mapped_data)

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
    parser_generate.add_argument('--template', action='store_true')
    parser_generate.add_argument('--manufacturer')
    parser_generate.add_argument('--model')
    parser_generate.add_argument('-o', '--output', help='Output definition CSV')
    parser_generate.add_argument('--protocol', default='modbusRTU')
    parser_generate.add_argument('--category', default='Inverter')
    parser_generate.add_argument('--forced-write', default='')
    parser_generate.add_argument('--address-offset', type=int, default=0, help='Address offset')

    # Validate
    parser_validate = subparsers.add_parser('validate', help='Validate existing definition file')
    parser_validate.add_argument('input_file', help='Definition CSV to validate')

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
    if not args.command:
        parser.print_help()
        return

    setup_logging(args.verbose)

    # Validate --pages
    pages_arg = getattr(args, 'pages', None)
    if pages_arg:
        ext = os.path.splitext(args.input_file)[1].lower()
        if ext != '.pdf':
            logging.warning("--pages is only applicable for PDF files. Ignoring.")
        else:
            try:
                [int(p.strip()) for p in pages_arg.split(',')]
            except ValueError:
                logging.error("Invalid format for --pages. Expected comma-separated integers.")
                sys.exit(1)

    # Manual validation for manufacturer and model
    template_mode = getattr(args, 'template', False) or getattr(args, 'forced_write', '') == 'TEMPLATE'
    if args.command in ['generate', 'run'] and not template_mode:
        if not getattr(args, 'manufacturer', None) or not getattr(args, 'model', None):
            logging.error("the following arguments are required: --manufacturer, --model")
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
