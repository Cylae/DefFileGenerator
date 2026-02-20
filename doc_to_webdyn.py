#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
import csv
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator, run_generator

def main():
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', help='Path to the documentation file (PDF, Excel, CSV)')
    parser.add_argument('--manufacturer', required=True, help='Manufacturer name')
    parser.add_argument('--model', required=True, help='Model name')
    parser.add_argument('-o', '--output', help='Output definition CSV filename')
    parser.add_argument('--protocol', default='modbusRTU', help='Protocol name')
    parser.add_argument('--category', default='Inverter', help='Device category')
    parser.add_argument('--sheet', help='Excel sheet name')
    parser.add_argument('--pages', help='PDF pages (comma separated)')
    parser.add_argument('--address-offset', type=int, default=0, help='Offset to subtract from addresses')
    parser.add_argument('-v', '--verbose', action='store_true', help='Verbose logging')

    args = parser.parse_args()

    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=log_level, format='%(levelname)s: %(message)s')

    logging.info(f"Processing {args.input_file} for {args.manufacturer} {args.model}")

    extractor = Extractor()
    ext = os.path.splitext(args.input_file)[1].lower()

    tables = []
    if ext == '.csv':
        # Special handling for CSV in doc_to_webdyn to maintain compatibility
        try:
            with open(args.input_file, 'r', encoding='utf-8-sig') as f:
                content = f.read(1024)
                f.seek(0)
                sniffer = csv.Sniffer()
                try:
                    dialect = sniffer.sniff(content, delimiters=";,")
                except csv.Error:
                    dialect = csv.excel
                    dialect.delimiter = ',' if ',' in content else ';'
                reader = csv.DictReader(f, dialect=dialect)
                rows = list(reader)
                if rows:
                    tables = [rows]
        except Exception as e:
            logging.error(f"Error reading CSV: {e}")
            sys.exit(1)
    elif ext in ['.xlsx', '.xls', '.xlsm']:
        tables = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        pages = [int(p.strip()) for p in args.pages.split(',')] if args.pages else None
        tables = extractor.extract_from_pdf(args.input_file, pages)
    else:
        logging.error(f"Unsupported file extension: {ext}")
        sys.exit(1)

    if not tables:
        logging.error("No data could be extracted.")
        sys.exit(1)

    mapped_data = extractor.map_and_clean(tables)

    if not mapped_data:
        logging.error("No registers could be identified.")
        sys.exit(1)

    # Determine output filename
    output_file = args.output
    if not output_file:
        safe_mfg = re.sub(r'[^a-zA-Z0-9]', '_', args.manufacturer).lower()
        safe_model = re.sub(r'[^a-zA-Z0-9]', '_', args.model).lower()
        output_file = f"{safe_mfg}_{safe_model}_definition.csv"

    # Use Generator via run_generator logic (we can use a temp file or programmatic call)
    # Since run_generator expects a file, we can either use a temp file or call process_rows directly
    # and then write the final CSV here.

    generator = Generator(address_offset=args.address_offset)
    processed_rows = generator.process_rows(mapped_data)

    try:
        with open(output_file, 'w', newline='', encoding='utf-8') as outfile:
            header_row = [args.protocol, args.category, args.manufacturer, args.model, '', '', '', '', '', '', '']
            writer = csv.writer(outfile, delimiter=';', lineterminator='\n')
            writer.writerow(header_row)

            for index, row in enumerate(processed_rows, start=1):
                data_row = [
                    str(index),
                    row['Info1'],
                    row['Info2'],
                    row['Info3'],
                    row['Info4'],
                    row['Name'],
                    row['Tag'],
                    row['CoefA'],
                    row['CoefB'],
                    row['Unit'],
                    row['Action']
                ]
                writer.writerow(data_row)
        logging.info(f"Definition file successfully generated at {output_file}")
    except Exception as e:
        logging.error(f"Error writing output: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
