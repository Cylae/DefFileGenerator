#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
import json
from DefFileGenerator.extractor import Extractor, peek_generator
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator

def _run_cli(argv=None):
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', nargs='?', help='Path to documentation (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer', help='Manufacturer name')
    parser.add_argument('--model', help='Model name')
    parser.add_argument('--template', action='store_true', help='Generate a template definition')
    parser.add_argument('-o', '--output', help='Output filename')
    parser.add_argument('--protocol', default='modbusRTU')
    parser.add_argument('--category', default='Inverter')
    parser.add_argument('--sheet', help='Excel sheet name')
    parser.add_argument('--pages', help='PDF pages (comma-separated integers)')
    parser.add_argument('--mapping', help='JSON mapping file')
    parser.add_argument('--address-offset', type=int, default=0)
    parser.add_argument('--forced-write', default='')
    parser.add_argument('-v', '--verbose', action='store_true')

    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format='%(levelname)s: %(message)s', force=True)

    if getattr(args, 'template', False):
        config = GeneratorConfig(output=args.output, template=True)
        run_generator(config)
        return

    input_file = args.input_file
    if not input_file or not os.path.exists(input_file):
        logging.error(f"Input file not found: {input_file}")
        sys.exit(1)

    ext = os.path.splitext(input_file)[1].lower()

    # Warn about mismatched options
    if args.pages and ext != '.pdf':
        logging.warning("--pages is only applicable for PDF files. Ignoring.")
    if args.sheet and ext not in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        logging.warning("--sheet is only applicable for Excel files. Ignoring.")

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

    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw = extractor.extract_from_excel(input_file, args.sheet)
    elif ext == '.pdf':
        raw = extractor.extract_from_pdf(input_file, args.pages)
    elif ext == '.csv':
        raw = extractor.extract_from_csv(input_file)
    elif ext == '.xml':
        raw = extractor.extract_from_xml(input_file)
    else:
        logging.error(f"Unsupported extension: {ext}")
        sys.exit(1)

    has_data, raw_peeked = peek_generator(raw)
    if not has_data:
        logging.error("No data extracted.")
        sys.exit(1)

    mapped = extractor.map_and_clean(raw_peeked, args.address_offset)
    has_regs, mapped_peeked = peek_generator(mapped)
    if not has_regs:
        logging.error("No registers extracted.")
        sys.exit(1)

    m_name = args.manufacturer or "Manufacturer"
    m_model = args.model or "Model"
    output_file = args.output or f"{re.sub(r'[^a-zA-Z0-9]', '_', m_name).lower()}_{re.sub(r'[^a-zA-Z0-9]', '_', m_model).lower()}_definition.csv"

    config = GeneratorConfig(
        input_file=input_file,
        output=output_file,
        manufacturer=m_name,
        model=m_model,
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        address_offset=0, # Already applied during extraction
        template=False
    )
    run_generator(config, input_data=mapped_peeked)

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

if __name__ == "__main__":
    main()
