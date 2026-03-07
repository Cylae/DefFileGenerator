#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
import csv
import io
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator

try:
    from defusedxml import ElementTree as ET
    HAS_DEFUSEDXML = True
except ImportError:
    HAS_DEFUSEDXML = False

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

    ext = os.path.splitext(args.input_file)[1].lower()
    extractor = Extractor()
    tables = []

    if ext == '.csv':
        tables = load_csv(args.input_file)
    elif ext in ['.xlsx', '.xls', '.xlsm']:
        tables = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.xml':
        tables = load_xml(args.input_file)
    elif ext == '.pdf':
        tables = extractor.extract_from_pdf(args.input_file)
    else:
        logging.error(f"Unsupported file extension: {ext}")
        sys.exit(1)

    if not tables:
        logging.error("No data could be extracted from the file.")
        sys.exit(1)

    # Use Extractor to map and clean the data
    mapped_data = extractor.map_and_clean(tables)

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

    # Use Generator to process and write output
    generator = Generator()
    processed_rows = generator.process_rows(mapped_data)

    config = GeneratorConfig(
        manufacturer=args.manufacturer,
        model=args.model,
        output=output_file,
        protocol=args.protocol,
        category=args.category
    )
    generator.write_output_csv(processed_rows, config)
    logging.info(f"Definition file successfully generated at {output_file}")

def load_csv(filepath):
    # Try different encodings and delimiters
    for encoding in ['utf-8-sig', 'utf-16', 'latin-1']:
        for delimiter in [',', ';', '\t']:
            try:
                with open(filepath, 'r', encoding=encoding) as f:
                    # Peek to check if it looks like CSV
                    sample = f.read(1024)
                    f.seek(0)
                    if delimiter not in sample:
                        continue
                    reader = csv.DictReader(f, delimiter=delimiter)
                    rows = list(reader)
                    if rows and len(reader.fieldnames) > 1:
                        return [rows]
            except Exception:
                continue
    return []

def load_xml(filepath):
    if not HAS_DEFUSEDXML:
        logging.error("defusedxml is not installed. XML parsing aborted for security.")
        return []

    try:
        import pandas as pd
        with open(filepath, 'rb') as f:
            xml_data = f.read()
        # Use defusedxml to parse first to ensure it's safe
        ET.fromstring(xml_data)
        # If it passes, use pandas for convenient table extraction via buffer
        df = pd.read_xml(io.BytesIO(xml_data))
        return [df.to_dict(orient='records')]
    except Exception as e:
        logging.error(f"Error loading XML file: {e}")
        return []

if __name__ == "__main__":
    main()
