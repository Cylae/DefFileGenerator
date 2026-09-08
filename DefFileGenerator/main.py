#!/usr/bin/env python3
"""
Primary Command Line Interface for WebdynSunPM Definition Tool.

Provides sub-commands (`run`, `extract`, `generate`, `validate`) to extract registers
from documentation files, convert them into WebdynSunPM format, and validate output files.
"""

import argparse
import csv
import json
import logging
import os
import re
import sys

# Ensure parent directory is in sys.path to support direct and packaged executions
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator, run_generator, GeneratorConfig, peek_generator

ALLOWED_EXTENSIONS = {'.pdf', '.xlsx', '.xlsm', '.xltx', '.xltm', '.csv', '.xml'}

def setup_logging(verbose=False, quiet=False):
    level = logging.DEBUG if verbose else (logging.WARNING if quiet else logging.INFO)
    logging.basicConfig(
        level=level,
        format='%(levelname)s: %(message)s',
        force=True
    )

def _get_default_output(manufacturer, model):
    m = re.sub(r'[^a-zA-Z0-9]', '_', manufacturer).lower()
    md = re.sub(r'[^a-zA-Z0-9]', '_', model).lower()
    return f"{m}_{md}_definition.csv"

def _perform_extraction(args):
    input_file = getattr(args, 'input_file', None)
    if not input_file:
        logging.error("Input file is required.")
        sys.exit(1)
    if not os.path.exists(input_file):
        logging.error(f"Input file not found: {input_file}")
        sys.exit(1)

    ext = os.path.splitext(input_file)[1].lower()
    if ext not in ALLOWED_EXTENSIONS:
        logging.error(
            f"Unsupported file type: {ext}. Supported formats are: .pdf, .xlsx, .xlsm, .xltx, .xltm, .csv, .xml"
        )
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
        logging.error(f"Unsupported file type: {ext}")
        sys.exit(1)

    has_data, raw_peeked = peek_generator(raw_data)
    if not has_data:
        logging.error("No data extracted.")
        sys.exit(1)

    mapped_gen = extractor.map_and_clean(raw_peeked, address_offset)
    has_regs, mapped_peeked = peek_generator(mapped_gen)
    if not has_regs:
        logging.error("No registers extracted.")
        sys.exit(1)

    return mapped_peeked

def extract_command(args):
    output = getattr(args, 'output', None)
    if output and os.path.exists(output) and not getattr(args, 'force', False):
        logging.error(f"Output file '{output}' already exists. Use --force to overwrite.")
        sys.exit(1)

    mapped_data = _perform_extraction(args)
    first, mapped_data_iter = peek_generator(mapped_data)
    if not first:
        logging.error("No registers extracted.")
        sys.exit(1)

    fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']

    if output:
        f = open(output, 'w', newline='', encoding='utf-8')
    else:
        f = sys.stdout

    try:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(mapped_data_iter)
    finally:
        if output:
            f.close()
            logging.info(f"Extraction complete. Saved to {output}")

def validate_command(args):
    input_file = getattr(args, 'input_file', None)
    if not input_file or not os.path.exists(input_file):
        logging.error(f"File not found: {input_file}")
        sys.exit(1)

    generator = Generator()
    strict = not getattr(args, 'lenient', False)
    if generator.validate_csv(input_file, strict=strict):
        logging.info(f"Validation successful: {input_file}")
    else:
        logging.error(f"Validation failed: {input_file}")
        sys.exit(1)

def generate_command(args):
    template = getattr(args, 'template', False)
    template_mode = getattr(args, 'template_mode', 'input')
    m_name = getattr(args, 'manufacturer', None)
    m_model = getattr(args, 'model', None)
    input_file = getattr(args, 'input_file', None)
    output = getattr(args, 'output', None)

    if not template:
        if not m_name or not m_model:
            logging.error("--manufacturer and --model are required for generate command.")
            sys.exit(1)
        if not input_file:
            logging.error("input_file is required for generate command.")
            sys.exit(1)
        if not os.path.exists(input_file):
            logging.error(f"Input file not found: {input_file}")
            sys.exit(1)

    if output and os.path.exists(output) and not getattr(args, 'force', False):
        logging.error(f"Output file '{output}' already exists. Use --force to overwrite.")
        sys.exit(1)

    config = GeneratorConfig(
        input_file=input_file,
        output=output,
        manufacturer=m_name or 'Manufacturer',
        model=m_model or 'Model',
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
    m_name = getattr(args, 'manufacturer', None)
    m_model = getattr(args, 'model', None)
    input_file = getattr(args, 'input_file', None)

    if not template:
        if not m_name or not m_model:
            logging.error("--manufacturer and --model are required for run command.")
            sys.exit(1)
        if not input_file:
            logging.error("input_file is required for run command.")
            sys.exit(1)
        if not os.path.exists(input_file):
            logging.error(f"Input file not found: {input_file}")
            sys.exit(1)

    mapped_data = None
    if not template:
        mapped_data = _perform_extraction(args)
        has_regs, mapped_data = peek_generator(mapped_data)
        if not has_regs:
            logging.error("No registers extracted.")
            sys.exit(1)

    output_file = getattr(args, 'output', None)
    if not output_file and not template:
        sanitized_mfg = re.sub(r'[^a-zA-Z0-9]', '_', m_name).lower()
        sanitized_model = re.sub(r'[^a-zA-Z0-9]', '_', m_model).lower()
        output_file = f"{sanitized_mfg}_{sanitized_model}_definition.csv"

    if output_file and os.path.exists(output_file) and not getattr(args, 'force', False):
        logging.error(f"Output file '{output_file}' already exists. Use --force to overwrite.")
        sys.exit(1)

    config = GeneratorConfig(
        input_file=input_file,
        output=output_file,
        manufacturer=m_name or 'Manufacturer',
        model=m_model or 'Model',
        protocol=getattr(args, 'protocol', 'modbusRTU'),
        category=getattr(args, 'category', 'Inverter'),
        forced_write=getattr(args, 'forced_write', ''),
        address_offset=0, # Already applied during extraction
        template=template,
        template_mode=getattr(args, 'template_mode', 'input')
    )
    run_generator(config, input_data=mapped_data)

    if output_file and not getattr(args, 'no_validate', False) and not template:
        generator = Generator()
        if generator.validate_csv(output_file):
            logging.info("Post-generation validation passed.")
        else:
            logging.error("Post-generation validation failed.")
            sys.exit(1)

def _run_cli(argv=None):
    if argv is None:
        argv = sys.argv[1:]
    else:
        if argv and (argv[0].endswith('main.py') or argv[0].endswith('doc_to_webdyn.py') or argv[0] == 'main.py' or argv[0] == 'doc_to_webdyn.py'):
            argv = argv[1:]

    parser = argparse.ArgumentParser(
        prog='deffilegen',
        description='WebdynSunPM Definition Tool',
        epilog="Examples:\n  deffilegen run input.xlsx --manufacturer Huawei --model SUN2000 -o huawei.csv\n  deffilegen extract doc.pdf -o registers.csv\n  deffilegen generate registers.csv --manufacturer SMA --model STP5000 -o sma.csv\n  deffilegen validate definition.csv\n\nExit codes:\n  0: Success\n  1: Execution error\n  2: Usage / argument error",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument('--version', action='version', version='deffilegen 0.2.1')
    parser.add_argument('-v', '--verbose', action='store_true', help='Verbose logging')
    parser.add_argument('-q', '--quiet', action='store_true', help='Quiet logging')

    subparsers = parser.add_subparsers(dest='command', help='Sub-commands')

    def add_common_flags(p):
        p.add_argument('-v', '--verbose', action='store_true', help='Verbose logging')
        p.add_argument('-q', '--quiet', action='store_true', help='Quiet logging')

    # Validate
    parser_validate = subparsers.add_parser('validate', help='Validate a WebdynSunPM definition file')
    add_common_flags(parser_validate)
    parser_validate.add_argument('input_file', help='Definition CSV to validate')
    parser_validate.add_argument('--lenient', action='store_true', help='Lenient validation (allow address overlaps)')

    # Extract Subparser
    parser_extract = subparsers.add_parser('extract', help='Extract registers from documentation')
    add_common_flags(parser_extract)
    parser_extract.add_argument('input_file', help='Source file (PDF/Excel/CSV/XML)')
    parser_extract.add_argument('-o', '--output', help='Output CSV')
    parser_extract.add_argument('--mapping', help='Mapping JSON')
    parser_extract.add_argument('--sheet', help='Excel sheet')
    parser_extract.add_argument('--pages', help='PDF pages')
    parser_extract.add_argument('--address-offset', type=int, default=0, help='Address offset')
    parser_extract.add_argument('--force', action='store_true', help='Force overwrite output file')

    # Generate
    parser_generate = subparsers.add_parser('generate', help='Generate definition from CSV')
    add_common_flags(parser_generate)
    parser_generate.add_argument('input_file', nargs='?', help='Input CSV')
    parser_generate.add_argument('-o', '--output', help='Output definition CSV')
    parser_generate.add_argument('--manufacturer', help='Manufacturer name')
    parser_generate.add_argument('--model', help='Model name')
    parser_generate.add_argument('--template', action='store_true', help='Generate template')
    parser_generate.add_argument('--template-mode', choices=['input', 'definition'], default='input')
    parser_generate.add_argument('--protocol', default='modbusRTU')
    parser_generate.add_argument('--category', default='Inverter')
    parser_generate.add_argument('--forced-write', default='')
    parser_generate.add_argument('--address-offset', type=int, default=0, help='Address offset')
    parser_generate.add_argument('--force', action='store_true', help='Force overwrite output file')

    # Run
    parser_run = subparsers.add_parser('run', help='Extract and Generate in one step')
    add_common_flags(parser_run)
    parser_run.add_argument('input_file', nargs='?', help='Source file (PDF/Excel/CSV/XML)')
    parser_run.add_argument('-o', '--output', help='Output definition CSV')
    parser_run.add_argument('--manufacturer', help='Manufacturer name')
    parser_run.add_argument('--model', help='Model name')
    parser_run.add_argument('--template', action='store_true', help='Generate template')
    parser_run.add_argument('--template-mode', choices=['input', 'definition'], default='input')
    parser_run.add_argument('--mapping', help='Mapping JSON')
    parser_run.add_argument('--sheet', help='Excel sheet')
    parser_run.add_argument('--pages', help='PDF pages')
    parser_run.add_argument('--protocol', default='modbusRTU')
    parser_run.add_argument('--category', default='Inverter')
    parser_run.add_argument('--forced-write', default='')
    parser_run.add_argument('--address-offset', type=int, default=0, help='Address offset')
    parser_run.add_argument('--force', action='store_true', help='Force overwrite output file')
    parser_run.add_argument('--no-validate', action='store_true', help='Skip post-generation validation')

    args = parser.parse_args(argv)

    if not args.command:
        parser.print_help()
        sys.exit(2)

    verbose = getattr(args, 'verbose', False)
    quiet = getattr(args, 'quiet', False)
    setup_logging(verbose=verbose, quiet=quiet)

    # Validate pages/sheet parameters depending on input file type
    input_file = getattr(args, 'input_file', None)
    if input_file:
        ext = os.path.splitext(input_file)[1].lower()
        if getattr(args, 'pages', None) and ext != '.pdf':
            logging.warning("--pages is only applicable for PDF files. Ignoring.")
        if getattr(args, 'sheet', None) and ext not in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
            logging.warning("--sheet is only applicable for Excel files. Ignoring.")

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
        raise e
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()
