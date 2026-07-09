#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
from DefFileGenerator.main import setup_logging, _perform_extraction
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator

def _run_cli():
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', nargs='?', help='Path to documentation (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer', help='Manufacturer name')
    parser.add_argument('--model', help='Model name')
    parser.add_argument('--template', action='store_true', help='Generate a template')
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
        config = GeneratorConfig(
            output=args.output,
            template=True,
            template_mode=args.template_mode
        )
        run_generator(config)
        return

    if not args.input_file:
        logging.error("input_file is required.")
        sys.exit(1)

    mapped_data = _perform_extraction(args)

    m_name = args.manufacturer or "Manufacturer"
    m_model = args.model or "Model"

    if args.output:
        output_file = args.output
    else:
        safe_mfg = re.sub(r'[^a-zA-Z0-9]', '_', m_name).lower()
        safe_model = re.sub(r'[^a-zA-Z0-9]', '_', m_model).lower()
        output_file = f"{safe_mfg}_{safe_model}_definition.csv"

    config = GeneratorConfig(
        input_file=args.input_file,
        output=output_file,
        manufacturer=m_name,
        model=m_model,
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        address_offset=0 # Already applied during extraction
    )
    run_generator(config, input_data=mapped_data)

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

if __name__ == "__main__":
    main()
