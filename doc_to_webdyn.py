#!/usr/bin/env python3
"""
Single-Step Documentation to WebdynSunPM Converter Interface.

Parses manufacturer register documentation (PDF, Excel, CSV, XML) and generates a
validated WebdynSunPM definition CSV file in a single step.
"""

import argparse
import json
import logging
import os
import re
import csv
import json
from DefFileGenerator.extractor import Extractor, peek_generator
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator

def _run_cli(args_list=None):
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', nargs='?', help='Path to documentation (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer', help='Manufacturer name')
    parser.add_argument('--model', help='Model name')
    parser.add_argument('--template', action='store_true', help='Generate a template definition')
    parser.add_argument('--template-mode', choices=['input', 'definition'], default='input')
    parser.add_argument('-o', '--output', help='Output filename')
    parser.add_argument('--protocol', default='modbusRTU')
    parser.add_argument('--category', default='Inverter')
    parser.add_argument('--sheet', help='Excel sheet name')
    parser.add_argument('--pages', help='PDF pages (comma-separated integers)')
    parser.add_argument('--mapping', help='JSON mapping file')
    parser.add_argument('--address-offset', type=int, default=0)
    parser.add_argument('--forced-write', default='')
    parser.add_argument('-v', '--verbose', action='store_true')

    args = parser.parse_args()
    setup_logging(args.verbose)

    if args.template:
        config = GeneratorConfig(output=args.output, template=True, template_mode=args.template_mode)
        run_generator(config)
        return

    input_file = getattr(args, 'input_file', None)
    if not input_file or not os.path.exists(input_file):
        logging.error(f"Input file not found: {input_file}")
        sys.exit(1)

    input_file = args.input_file
    ext = os.path.splitext(input_file)[1].lower()

    # Warn about mismatched options
    if args.pages and ext != '.pdf':
        logging.warning("--pages is only applicable for PDF files. Ignoring.")
    if args.sheet and ext not in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        logging.warning("--sheet is only applicable for Excel files. Ignoring.")

    mapping = {}
    if args.mapping:
        if not os.path.exists(args.mapping):
            logging.error(f"Mapping file not found: {args.mapping}")
            sys.exit(1)
        try:
            with open(args.mapping, 'r') as f:
                mapping = json.load(f)
        except Exception as e:
            logging.error(f"Error reading mapping file: {e}")
            sys.exit(1)

    extractor = Extractor(mapping)
    ext = os.path.splitext(args.input_file)[1].lower()

    pages = args.pages
    if pages and ext == '.pdf':
        try:
            # Extractor expects pages as comma-separated string or list of ints.
            # Our current extractor.extract_from_pdf handles string.
            pass
        except ValueError:
            logging.error("Invalid format for --pages.")
            sys.exit(1)

    sheet_arg = getattr(args, 'sheet', None)
    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']: raw = extractor.extract_from_excel(args.input_file, sheet_arg)
    elif ext == '.pdf': raw = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv': raw = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml': raw = extractor.extract_from_xml(args.input_file)
    else: logging.error(f"Unsupported extension: {ext}"); sys.exit(1)

    has_data, raw = peek_generator(raw)
    if not has_data: logging.error("No data extracted."); sys.exit(1)

    mapped = extractor.map_and_clean(raw, args.address_offset)
    has_registers, mapped = peek_generator(mapped)
    if not has_registers: logging.error("No registers extracted."); sys.exit(1)

    m_name = args.manufacturer or "Manufacturer"
    m_model = args.model or "Model"

    if args.output:
        output_file = args.output
    else:
        m_name_clean = re.sub(r'[^a-zA-Z0-9]', '_', m_name).lower()
        m_model_clean = re.sub(r'[^a-zA-Z0-9]', '_', m_model).lower()
        output_file = f"{m_name_clean}_{m_model_clean}_definition.csv"

    config = GeneratorConfig(
        input_file=args.input_file,
        output=output_file,
        manufacturer=m_name,
        model=m_model,
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        address_offset=0, # Already applied during extraction
        template=getattr(args, 'template', False)
    )
    run_generator(config, input_data=mapped)

def setup_logging(verbose):
    logging.basicConfig(level=logging.DEBUG if verbose else logging.INFO, format='%(levelname)s: %(message)s', force=True)

def main(args=None):
    try:
        _run_cli(args)
    except KeyboardInterrupt:
        sys.exit(130)
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        # traceback.print_exc() # For deep debugging
        sys.exit(1)


if __name__ == "__main__":
    main()