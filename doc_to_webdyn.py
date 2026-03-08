#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator

def main():
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', help='Path to manufacturer documentation (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer', required=True, help='Manufacturer name')
    parser.add_argument('--model', required=True, help='Model name')
    parser.add_argument('-o', '--output', help='Output filename')
    parser.add_argument('--protocol', default='modbusRTU', help='Protocol name')
    parser.add_argument('--category', default='Inverter', help='Device category')
    parser.add_argument('--sheet', help='Excel sheet name')
    parser.add_argument('-v', '--verbose', action='store_true', help='Verbose output')
    parser.add_argument('--address-offset', type=int, default=0, help='Address offset')

    args = parser.parse_args()

    # Configure logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=log_level, format='%(levelname)s: %(message)s', force=True)

    logging.info(f"Processing {args.input_file} for {args.manufacturer} {args.model}")

    extractor = Extractor()
    ext = os.path.splitext(args.input_file)[1].lower()

    if ext in ['.xlsx', '.xls', '.xlsm']:
        raw_data = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        raw_data = extractor.extract_from_pdf(args.input_file)
    elif ext == '.csv':
        raw_data = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml':
        raw_data = extractor.extract_from_xml(args.input_file)
    else:
        logging.error(f"Unsupported extension: {ext}")
        sys.exit(1)

    if not raw_data:
        logging.error("No data extracted.")
        sys.exit(1)

    mapped_data = extractor.map_and_clean(raw_data)
    if not mapped_data:
        logging.error("No registers found.")
        sys.exit(1)

    logging.info(f"Extracted {len(mapped_data)} registers.")

    # Determine output filename
    output_file = args.output
    if not output_file:
        safe_mfg = re.sub(r'[^a-zA-Z0-9]', '_', args.manufacturer).lower()
        safe_model = re.sub(r'[^a-zA-Z0-9]', '_', args.model).lower()
        output_file = f"{safe_mfg}_{safe_model}_definition.csv"

    # Process and write
    generator = Generator()
    processed_rows = generator.process_rows(mapped_data, args.address_offset)

    config = GeneratorConfig(
        manufacturer=args.manufacturer,
        model=args.model,
        protocol=args.protocol,
        category=args.category,
        address_offset=args.address_offset
    )
    generator.write_output_csv(output_file, processed_rows, config)

if __name__ == "__main__":
    main()
