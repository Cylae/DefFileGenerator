#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator, GeneratorConfig

def main():
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', help='Path to the manufacturer documentation file (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer', required=True, help='Manufacturer name')
    parser.add_argument('--model', required=True, help='Model name')
    parser.add_argument('-o', '--output', help='Output filename (default: auto-generated)')
    parser.add_argument('--protocol', default='modbusRTU', help='Protocol name (default: modbusRTU)')
    parser.add_argument('--category', default='Inverter', help='Device category (default: Inverter)')
    parser.add_argument('--sheet', help='Excel sheet name (processes all if not specified)')
    parser.add_argument('-v', '--verbose', action='store_true', help='Show detailed processing information')

    args = parser.parse_args()

    # Configure logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=log_level, format='%(levelname)s: %(message)s', force=True)

    logging.info(f"Processing {args.input_file} for {args.manufacturer} {args.model}")

    extractor = Extractor()
    ext = os.path.splitext(args.input_file)[1].lower()
    raw_data = []

    if ext == '.csv':
        raw_data = extractor.extract_from_csv(args.input_file)
    elif ext in ['.xlsx', '.xls']:
        raw_data = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.xml':
        raw_data = extractor.extract_from_xml(args.input_file)
    elif ext == '.pdf':
        raw_data = extractor.extract_from_pdf(args.input_file)
    else:
        logging.error(f"Unsupported file extension: {ext}")
        sys.exit(1)

    if not raw_data:
        logging.error("No data could be extracted from the file.")
        sys.exit(1)

    mapped_data = extractor.map_and_clean(raw_data)

    if not mapped_data:
        logging.error("No registers could be extracted from the tables.")
        sys.exit(1)

    logging.info(f"Successfully extracted {len(mapped_data)} registers.")

    # Use DefFileGenerator logic to process and validate
    generator = Generator()
    processed_rows = generator.process_rows(mapped_data)

    # Determine output filename
    output_file = args.output
    if not output_file:
        safe_mfg = re.sub(r'[^a-zA-Z0-9]', '_', args.manufacturer).lower()
        safe_model = re.sub(r'[^a-zA-Z0-9]', '_', args.model).lower()
        output_file = f"{safe_mfg}_{safe_model}_definition.csv"

    # Write Output in WebdynSunPM format
    try:
        with open(output_file, 'w', newline='', encoding='utf-8') as outfile:
            Generator.write_output_csv(
                outfile,
                args.protocol,
                args.category,
                args.manufacturer,
                args.model,
                '', # ForcedWriteCode
                processed_rows
            )
        logging.info(f"Definition file successfully generated at {output_file}")
    except Exception as e:
        logging.error(f"Error writing output file: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
