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

logger = logging.getLogger('DefFileGenerator.main')

def setup_logging(verbose=False):
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format='%(levelname)s: %(message)s',
        force=True
    )

def _add_common_args(parser, include_input=True):
    if include_input:
        parser.add_argument('input_file', nargs='?', help='Source/Input file')
    parser.add_argument('--manufacturer', help='Manufacturer name')
    parser.add_argument('--model', help='Model name')
    parser.add_argument('-o', '--output', help='Output filename')
    parser.add_argument('--protocol', default='modbusRTU')
    parser.add_argument('--category', default='Inverter')
    parser.add_argument('--address-offset', type=int, default=0, help='Address offset')
    parser.add_argument('--forced-write', default='', help='Webdyn forced write config')

def _perform_extraction(args):
    input_file = getattr(args, 'input_file', None)
    if not input_file or not os.path.exists(input_file):
        logger.error(f"Input file not found: {input_file}")
        sys.exit(1)

    mapping = {}
    if getattr(args, 'mapping', None) and os.path.exists(args.mapping):
        try:
            with open(args.mapping, 'r') as f:
                mapping = json.load(f)
        except Exception as e:
            logger.error(f"Error reading mapping file: {e}")
            sys.exit(1)

    extractor = Extractor(mapping)
    ext = os.path.splitext(input_file)[1].lower()

    if ext in ['.xlsx', '.xlsm']: raw = extractor.extract_from_excel(input_file, getattr(args, 'sheet', None))
    elif ext == '.pdf': raw = extractor.extract_from_pdf(input_file, getattr(args, 'pages', None))
    elif ext == '.csv': raw = extractor.extract_from_csv(input_file)
    elif ext == '.xml': raw = extractor.extract_from_xml(input_file)
    else:
        logger.error(f"Unsupported extension: {ext}")
        sys.exit(1)

    has_data, raw_peeked = peek_generator(raw)
    if not has_data:
        logger.error("No data extracted.")
        sys.exit(1)

    mapped_gen = extractor.map_and_clean(raw_peeked, getattr(args, 'address_offset', 0))
    has_regs, mapped_peeked = peek_generator(mapped_gen)
    if not has_regs:
        logger.error("No registers extracted.")
        sys.exit(1)

    return mapped_peeked

def extract_command(args):
    mapped_data = _perform_extraction(args)
    output = args.output
    f = open(output, 'w', newline='', encoding='utf-8') if output else sys.stdout
    fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
    writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
    writer.writeheader()
    writer.writerows(mapped_data)
    if output:
        f.close()
        logger.info(f"Extraction complete. Saved to {output}")

def validate_command(args):
    if Generator().validate_csv(args.input_file, strict_overlap=getattr(args, 'strict_overlap', False)):
        logger.info(f"Validation successful: {args.input_file}")
    else:
        logger.error(f"Validation failed: {args.input_file}")
        sys.exit(1)

def generate_command(args):
    config = GeneratorConfig(
        input_file=args.input_file,
        output=args.output,
        manufacturer=args.manufacturer or 'Manufacturer',
        model=args.model or 'Model',
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        address_offset=args.address_offset,
        template=getattr(args, 'template', False),
        template_mode=getattr(args, 'template_mode', 'input')
    )
    run_generator(config)

def run_command(args):
    template = getattr(args, 'template', False)
    mapped_data = None
    if not template:
        if not args.manufacturer or not args.model:
            logger.error("--manufacturer and --model are required for run.")
            sys.exit(1)
        mapped_data = _perform_extraction(args)

    config = GeneratorConfig(
        input_file=args.input_file,
        output=args.output,
        manufacturer=args.manufacturer or 'Manufacturer',
        model=args.model or 'Model',
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        address_offset=0, # Already applied during extraction
        template=template,
        template_mode=getattr(args, 'template_mode', 'input')
    )
    run_generator(config, input_data=mapped_data)

def _run_cli():
    parser = argparse.ArgumentParser(description='WebdynSunPM Definition Tool')
    parser.add_argument('-v', '--verbose', action='store_true', help='Verbose logging')
    subparsers = parser.add_subparsers(dest='command', help='Sub-commands')

    # Validate
    parser_validate = subparsers.add_parser('validate', help='Validate a definition file')
    parser_validate.add_argument('input_file', help='Definition CSV to validate')
    parser_validate.add_argument('--strict-overlap', action='store_true', help='Treat address overlaps as fatal errors')

    # Extract
    parser_extract = subparsers.add_parser('extract', help='Extract registers from documentation')
    _add_common_args(parser_extract)
    parser_extract.add_argument('--mapping', help='Mapping JSON')
    parser_extract.add_argument('--sheet', help='Excel sheet')
    parser_extract.add_argument('--pages', help='PDF pages')

    # Generate
    parser_generate = subparsers.add_parser('generate', help='Generate definition from CSV')
    _add_common_args(parser_generate)
    parser_generate.add_argument('--template', action='store_true')
    parser_generate.add_argument('--template-mode', choices=['input', 'definition'], default='input')

    # Run
    parser_run = subparsers.add_parser('run', help='Extract and Generate in one step')
    _add_common_args(parser_run)
    parser_run.add_argument('--template', action='store_true')
    parser_run.add_argument('--template-mode', choices=['input', 'definition'], default='input')
    parser_run.add_argument('--mapping', help='Mapping JSON')
    parser_run.add_argument('--sheet', help='Excel sheet')
    parser_run.add_argument('--pages', help='PDF pages')

    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        return

    setup_logging(args.verbose)

    if args.command == 'extract':
        extract_command(args)
    elif args.command == 'validate':
        validate_command(args)
    elif args.command == 'generate':
        if not getattr(args, 'template', False) and not args.input_file:
            logger.error("input_file is required for generate unless --template is used.")
            sys.exit(1)
        generate_command(args)
    elif args.command == 'run':
        run_command(args)

def main():
    try:
        _run_cli()
    except KeyboardInterrupt:
        sys.exit(130)
    except Exception as e:
        logger.error(f"Unexpected error: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()
