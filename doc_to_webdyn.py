#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
import csv
import io

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    from defusedxml import ElementTree as ET
    HAS_DEFUSEDXML = True
except ImportError:
    HAS_DEFUSEDXML = False

from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator

def load_xml(filepath):
    if not HAS_DEFUSEDXML:
        logging.error("defusedxml is required for secure XML processing. Please install it.")
        return []
    try:
        if HAS_PANDAS:
            with open(filepath, 'rb') as f:
                content = f.read()
            # pandas read_xml doesn't use defusedxml directly but we can parse it first
            # to check for XXE if needed, or use pandas with caution.
            # Project memory says: "XML parsing in doc_to_webdyn.py leverages defusedxml
            # to validate inputs against XXE vulnerabilities before construction of DataFrames"
            ET.fromstring(content) # This will raise an error if it's malicious
            df = pd.read_xml(io.BytesIO(content))
            return [df.to_dict('records')]
        else:
            tree = ET.parse(filepath)
            root = tree.getroot()
            # Heuristic XML to dict conversion
            data = []
            for child in root:
                row = {}
                for subchild in child:
                    row[subchild.tag] = subchild.text
                data.append(row)
            return [data]
    except Exception as e:
        logging.error(f"Error loading XML file: {e}")
        return []

def load_csv(filepath):
    # Extractor doesn't have a load_csv method, so we handle it here or add it there.
    # Centralizing in doc_to_webdyn for simplicity as Extractor focuses on Doc (PDF/Excel)
    all_data = []
    try:
        with open(filepath, 'rb') as f:
            raw = f.read(4)
            encoding = 'utf-16' if raw.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'

        with open(filepath, 'r', encoding=encoding) as f:
            content = f.read(4096)
            f.seek(0)
            try:
                sniffer = csv.Sniffer()
                dialect = sniffer.sniff(content, delimiters=";,")
            except csv.Error:
                dialect = csv.excel
                dialect.delimiter = ','

            reader = csv.DictReader(f, dialect=dialect)
            if reader.fieldnames:
                reader.fieldnames = [n.strip() for n in reader.fieldnames]
            for row in reader:
                all_data.append(row)
    except Exception as e:
        logging.error(f"Error loading CSV file: {e}")
    return [all_data]

def main():
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', help='Manufacturer documentation (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer', required=True, help='Manufacturer name')
    parser.add_argument('--model', required=True, help='Model name')
    parser.add_argument('-o', '--output', help='Output filename')
    parser.add_argument('--protocol', default='modbusRTU', help='Protocol (default: modbusRTU)')
    parser.add_argument('--category', default='Inverter', help='Category (default: Inverter)')
    parser.add_argument('--sheet', help='Excel sheet name')
    parser.add_argument('-v', '--verbose', action='store_true', help='Verbose logging')

    args = parser.parse_args()

    # Configure logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=log_level, format='%(levelname)s: %(message)s', force=True)

    logging.info(f"Processing {args.input_file} for {args.manufacturer} {args.model}")

    ext = os.path.splitext(args.input_file)[1].lower()
    extractor = Extractor()
    raw_data = []

    if ext == '.csv':
        raw_data = load_csv(args.input_file)
    elif ext in ['.xlsx', '.xls']:
        raw_data = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.xml':
        raw_data = load_xml(args.input_file)
    elif ext == '.pdf':
        raw_data = extractor.extract_from_pdf(args.input_file)
    else:
        logging.error(f"Unsupported file extension: {ext}")
        sys.exit(1)

    if not raw_data:
        logging.error("No data could be extracted.")
        sys.exit(1)

    mapped_data = extractor.map_and_clean(raw_data)
    if not mapped_data:
        logging.error("No registers could be extracted from the tables.")
        sys.exit(1)

    logging.info(f"Successfully extracted {len(mapped_data)} registers.")

    generator = Generator()
    processed_rows = generator.process_rows(mapped_data)

    output_file = args.output
    if not output_file:
        safe_mfg = re.sub(r'[^a-zA-Z0-9]', '_', args.manufacturer).lower()
        safe_model = re.sub(r'[^a-zA-Z0-9]', '_', args.model).lower()
        output_file = f"{safe_mfg}_{safe_model}_definition.csv"

    config = GeneratorConfig(
        manufacturer=args.manufacturer,
        model=args.model,
        protocol=args.protocol,
        category=args.category
    )

    generator.write_output_csv(output_file, config, processed_rows)
    logging.info(f"Definition file successfully generated at {output_file}")

if __name__ == "__main__":
    main()
