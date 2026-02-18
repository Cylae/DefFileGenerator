#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
import csv
import tempfile

from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import run_generator

def main():
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', help='Path to the manufacturer documentation file (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer', required=True, help='Manufacturer name')
    parser.add_argument('--model', required=True, help='Model name')
    parser.add_argument('-o', '--output', help='Output filename (default: auto-generated)')
    parser.add_argument('--protocol', default='modbusRTU', help='Protocol name (default: modbusRTU)')
    parser.add_argument('--category', default='Inverter', help='Device category (default: Inverter)')
    parser.add_argument('--sheet', help='Excel sheet name (processes all if not specified)')
    parser.add_argument('--address-offset', type=int, default=0, help='Value to subtract from all addresses (default: 0)')
    parser.add_argument('-v', '--verbose', action='store_true', help='Show detailed processing information')

    args = parser.parse_args()

    # Configure logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=log_level, format='%(levelname)s: %(message)s')

    logging.info(f"Processing {args.input_file} for {args.manufacturer} {args.model}")

    ext = os.path.splitext(args.input_file)[1].lower()
    extractor = Extractor()
    raw_data = []

    # Select loading method based on extension
    if ext in ['.xlsx', '.xls']:
        raw_data = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        raw_data = extractor.extract_from_pdf(args.input_file)
    elif ext == '.csv':
        # Simple CSV loading
        for delimiter in [',', ';', '\t']:
            try:
                with open(args.input_file, 'r', encoding='utf-8-sig') as f:
                    # Use Sniffer to improve detection if possible
                    content = f.read(2048)
                    f.seek(0)
                    try:
                        dialect = csv.Sniffer().sniff(content, delimiters=[',', ';', '\t'])
                        delimiter = dialect.delimiter
                    except Exception:
                        pass # use loop delimiter

                    reader = csv.DictReader(f, delimiter=delimiter)
                    rows = list(reader)
                    if rows and len(reader.fieldnames) > 1:
                        raw_data = rows
                        break
            except Exception:
                continue
    elif ext == '.xml':
        # XML loading (using pandas if available as in original)
        try:
            import pandas as pd
            df = pd.read_xml(args.input_file)
            raw_data = df.to_dict(orient='records')
        except ImportError:
            logging.error("pandas is required for XML processing.")
            sys.exit(1)
        except Exception as e:
            logging.error(f"Error loading XML: {e}")
            sys.exit(1)
    else:
        logging.error(f"Unsupported extension: {ext}")
        sys.exit(1)

    if not raw_data:
        logging.error("No data could be extracted from the file.")
        sys.exit(1)

    mapped_data = extractor.map_and_clean(raw_data)

    if not mapped_data:
        logging.error("No registers could be extracted from the tables.")
        sys.exit(1)

    logging.info(f"Successfully extracted {len(mapped_data)} registers.")

    # Determine output filename
    output_file = args.output
    if not output_file:
        safe_mfg = re.sub(r'[^a-zA-Z0-9]', '_', args.manufacturer).lower()
        safe_model = re.sub(r'[^a-zA-Z0-9]', '_', args.model).lower()
        output_file = f"{safe_mfg}_{safe_model}_definition.csv"

    # Use a temporary file for run_generator
    with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as tf:
        temp_csv = tf.name
        fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
        writer = csv.DictWriter(tf, fieldnames=fieldnames, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(mapped_data)

    try:
        run_generator(
            input_file=temp_csv,
            output=output_file,
            manufacturer=args.manufacturer,
            model=args.model,
            protocol=args.protocol,
            category=args.category,
            address_offset=args.address_offset
        )
        logging.info(f"Definition file successfully generated at {output_file}")
    finally:
        if os.path.exists(temp_csv):
            os.remove(temp_csv)

if __name__ == "__main__":
    main()
