#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import csv
import json
import re

# Ensure parent directory is in sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

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
        logging.error("Input file is required.")
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

    has_data, raw_data_peeked = peek_generator(raw_data)
    if not has_data:
        logging.error("No registers extracted.")
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
    if not generator.validate_csv(args.input_file):
        logging.error(f"Validation failed for {args.input_file}")
        sys.exit(1)
    logging.info(f"Validation successful for {args.input_file}")

def generate_command(args):
    template = getattr(args, 'template', False)
    template_mode = getattr(args, 'template_mode', 'input')
    input_file = getattr(args, 'input_file', None)

    if template and input_file == 'definition':
        template_mode = 'definition'
        input_file = None

    output = getattr(args, 'output', None)
    if not output and not template:
        mfg_clean = re.sub(r'[^a-zA-Z0-9_]', '', args.manufacturer.lower())
        model_clean = re.sub(r'[^a-zA-Z0-9_]', '', args.model.lower())
        output = f"{mfg_clean}_{model_clean}_definition.csv"

    config = GeneratorConfig(
        input_file=input_file,
        output=output,
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

def run_command(args):
    template = getattr(args, 'template', False)
    mapped_data = None
    if not template:
        mapped_data = _perform_extraction(args)

    output = getattr(args, 'output', None)
    if not output and not template:
        mfg_clean = re.sub(r'[^a-zA-Z0-9_]', '', args.manufacturer.lower())
        model_clean = re.sub(r'[^a-zA-Z0-9_]', '', args.model.lower())
        output = f"{mfg_clean}_{model_clean}_definition.csv"

    config = GeneratorConfig(
        input_file=getattr(args, 'input_file', None),
        output=output,
        manufacturer=getattr(args, 'manufacturer', 'Manufacturer'),
        model=getattr(args, 'model', 'Model'),
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        address_offset=0,  # Offset applied in _perform_extraction
        template=template,
        template_mode=getattr(args, 'template_mode', 'input')
    )
    run_generator(config, input_data=mapped_data)

def _run_cli(argv=None):
    if argv is None:
        argv = sys.argv[1:]

    # Strip script name if it matches 'main.py' or 'doc_to_webdyn.py'
    if argv and (argv[0].endswith('main.py') or argv[0].endswith('doc_to_webdyn.py')):
        argv = argv[1:]

    parser = argparse.ArgumentParser(description='WebdynSunPM Definition Tool')
    parser.add_argument('-v', '--verbose', action='store_true', help='Verbose logging')
    subparsers = parser.add_subparsers(dest='command', help='Sub-commands')

    # Validate Subparser
    parser_validate = subparsers.add_parser('validate', help='Validate a definition file')
    parser_validate.add_argument('input_file', help='Definition CSV to validate')

    # Extract Subparser
    parser_extract = subparsers.add_parser('extract', help='Extract registers from documentation')
    parser_extract.add_argument('input_file', help='Source file (PDF/Excel/CSV/XML)')
    parser_extract.add_argument('-o', '--output', help='Output CSV')
    parser_extract.add_argument('--mapping', help='Mapping JSON')
    parser_extract.add_argument('--sheet', help='Excel sheet')
    parser_extract.add_argument('--pages', help='PDF pages')
    parser_extract.add_argument('--address-offset', type=int, default=0, help='Address offset')

    # Generate Subparser
    parser_generate = subparsers.add_parser('generate', help='Generate definition from CSV')
    parser_generate.add_argument('input_file', nargs='?', help='Input CSV')
    parser_generate.add_argument('--manufacturer')
    parser_generate.add_argument('--model')
    parser_generate.add_argument('-o', '--output', help='Output definition CSV')
    parser_generate.add_argument('--protocol', default='modbusRTU')
    parser_generate.add_argument('--category', default='Inverter')
    parser_generate.add_argument('--forced-write', default='')
    parser_generate.add_argument('--template', action='store_true', help='Generate sample template')
    parser_generate.add_argument('--template-mode', choices=['input', 'definition'], default='input', help='Template mode')
    parser_generate.add_argument('--address-offset', type=int, default=0, help='Address offset')

    # Run Subparser
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
    parser_run.add_argument('--template', action='store_true', help='Generate sample template')
    parser_run.add_argument('--template-mode', choices=['input', 'definition'], default='input', help='Template mode')
    parser_run.add_argument('--address-offset', type=int, default=0, help='Address offset')

    args = parser.parse_args(argv)

    if not args.command:
        parser.print_help()
        sys.exit(0)

    setup_logging(args.verbose)

    # Validate --pages and --sheet warning
    pages_arg = getattr(args, 'pages', None)
    input_file = getattr(args, 'input_file', None)
    if pages_arg and input_file:
        ext = os.path.splitext(input_file)[1].lower()
        if ext != '.pdf':
            logging.warning("--pages is only applicable for PDF files. Ignoring.")

    sheet_arg = getattr(args, 'sheet', None)
    if sheet_arg and input_file:
        ext = os.path.splitext(input_file)[1].lower()
        if ext not in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
            logging.warning("--sheet is only applicable for Excel files. Ignoring.")

    # Validate required arguments for generate and run
    if args.command in ['generate', 'run'] and not getattr(args, 'template', False):
        if not getattr(args, 'manufacturer', None) or not getattr(args, 'model', None):
            logging.error(f"--manufacturer and --model are required for {args.command}.")
            sys.exit(1)
        if not getattr(args, 'input_file', None):
            logging.error(f"input_file is required for {args.command}.")
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
    except SystemExit as e:
        sys.exit(e.code)
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()
