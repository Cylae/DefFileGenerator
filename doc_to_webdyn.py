#!/usr/bin/env python3
import argparse
import json
import logging
import os
import re
import sys

from DefFileGenerator.def_gen import GeneratorConfig, run_generator
from DefFileGenerator.extractor import Extractor, peek_generator


def _run_cli(argv=None):
    if argv is None:
        argv = sys.argv[1:]
    else:
        # Strip script name if present as first element
        if argv and (
            argv[0].endswith('main.py')
            or argv[0].endswith('doc_to_webdyn.py')
            or argv[0] == 'main.py'
            or argv[0] == 'doc_to_webdyn.py'
        ):
            argv = argv[1:]

    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', nargs='?', help='Path to documentation (PDF/Excel/CSV/XML)')
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
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format='%(levelname)s: %(message)s',
        force=True,
    )

    if args.template:
        config = GeneratorConfig(output=args.output, template=True)
        run_generator(config)
        return

    if not args.input_file:
        logging.error("Input file is required.")
        sys.exit(1)

    if not os.path.exists(args.input_file):
        logging.error(f"Input file not found: {args.input_file}")
        sys.exit(1)

    input_file = args.input_file
    ext = os.path.splitext(input_file)[1].lower()

    mapping = {}
    if args.mapping:
        try:
            with open(args.mapping) as f:
                mapping = json.load(f)
        except (OSError, ValueError) as e:
            logging.error(f"Error reading mapping file: {e}")
            sys.exit(1)

    extractor = Extractor(mapping)

    pages = getattr(args, 'pages', None)
    if pages and ext == '.pdf':
        try:
            pages = [int(p.strip()) for p in pages.split(',')]
        except ValueError:
            logging.error("Invalid format for --pages. Expected comma-separated integers.")
            sys.exit(1)

    sheet_arg = getattr(args, 'sheet', None)
    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw = extractor.extract_from_excel(args.input_file, sheet_arg)
    elif ext == '.pdf':
        raw = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv':
        raw = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml':
        raw = extractor.extract_from_xml(args.input_file)
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
    output_file = args.output
    if not output_file:
        clean_mfg = re.sub(r'[^a-zA-Z0-9]', '_', m_name).lower()
        clean_model = re.sub(r'[^a-zA-Z0-9]', '_', m_model).lower()
        output_file = f"{clean_mfg}_{clean_model}_definition.csv"

    config = GeneratorConfig(
        input_file=args.input_file,
        output=output_file,
        manufacturer=m_name,
        model=m_model,
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        address_offset=0
    )
    run_generator(config, input_data=mapped_peeked)

def main(args=None):
    try:
        _run_cli(args)
    except KeyboardInterrupt:
        sys.exit(130)
    except SystemExit as e:
        sys.exit(e.code)
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        import traceback
        logging.debug(traceback.format_exc())
        sys.exit(1)


if __name__ == "__main__":
    main()