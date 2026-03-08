#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
from DefFileGenerator.main import _perform_extraction
from DefFileGenerator.def_gen import GeneratorConfig, Generator

def main():
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', help='Path to manufacturer documentation (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer', required=True, help='Manufacturer name')
    parser.add_argument('--model', required=True, help='Model name')
    parser.add_argument('-o', '--output', help='Output filename')
    parser.add_argument('--protocol', default='modbusRTU', help='Protocol name')
    parser.add_argument('--category', default='Inverter', help='Device category')
    parser.add_argument('--sheet', help='Excel sheet name')
    parser.add_argument('--pages', help='PDF pages (comma separated)')
    parser.add_argument('--address-offset', type=int, default=0, help='Register address offset')
    parser.add_argument('-v', '--verbose', action='store_true', help='Verbose logging')

    args = parser.parse_args()

    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=log_level, format='%(levelname)s: %(message)s', force=True)

    logging.info(f"Processing {args.input_file} for {args.manufacturer} {args.model}")

    # Use central extraction logic
    # _perform_extraction expects an args object with: input_file, mapping, sheet, pages
    # We add 'mapping' to args for compatibility
    args.mapping = None
    mapped_data = _perform_extraction(args)

    if not mapped_data:
        logging.error("No registers could be extracted.")
        sys.exit(1)

    logging.info(f"Extracted {len(mapped_data)} register(s).")

    # Use central generation logic
    config = GeneratorConfig(
        output=args.output,
        manufacturer=args.manufacturer,
        model=args.model,
        protocol=args.protocol,
        category=args.category,
        address_offset=args.address_offset
    )

    if not config.output:
        safe_mfg = re.sub(r'[^a-zA-Z0-9]', '_', args.manufacturer).lower()
        safe_model = re.sub(r'[^a-zA-Z0-9]', '_', args.model).lower()
        config.output = f"{safe_mfg}_{safe_model}_definition.csv"

    generator = Generator()
    processed_rows = generator.process_rows(mapped_data, address_offset=config.address_offset)
    generator.write_output_csv(config.output, processed_rows, config)

    logging.info(f"Definition file successfully generated at {config.output}")

if __name__ == "__main__":
    main()
