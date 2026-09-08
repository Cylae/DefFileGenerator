#!/usr/bin/env python3
"""
generate_webdyn_def.py

A robust programmatic wrapper and command-line demonstration tool for
end-to-end extraction, generation, and validation of WebdynSunPM definition files.
Uses the Extractor and Generator classes from the DefFileGenerator package.
"""

import argparse
import sys
import os
import logging
import re
import csv
import json
from DefFileGenerator.extractor import Extractor, peek_generator
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator

def _run_cli(argv=None):
    if argv is None:
        argv = sys.argv[1:]

    parser = argparse.ArgumentParser(
        description='Generate and Validate WebdynSunPM Definition File (.csv) from Manufacturer Documentation'
    )
    # Positional Arguments
    parser.add_argument('input_file', help='Path to documentation (PDF, Excel, CSV, XML)')
    parser.add_argument('output_file', help='Path to output WebdynSunPM definition CSV file')

    # Required Named Arguments
    parser.add_argument('--manufacturer', required=True, help='Manufacturer name')
    parser.add_argument('--model', required=True, help='Model name')

    # Optional Named Arguments
    parser.add_argument('--protocol', default='modbusRTU', help='Protocol name (default: modbusRTU)')
    parser.add_argument('--category', default='Inverter', help='Device category (default: Inverter)')
    parser.add_argument('--address-offset', type=int, default=0, help='Address offset to apply to registers (default: 0)')
    parser.add_argument('--sheet', help='Excel sheet name to process')
    parser.add_argument('--pages', help='PDF pages (comma-separated integers)')
    parser.add_argument('--mapping', help='JSON mapping file to customize column headers')
    parser.add_argument('--forced-write', default='', help='Force specific register write options')
    parser.add_argument('-v', '--verbose', action='store_true', help='Enable verbose/debug logging')

    args = parser.parse_args(argv)

    # Configure logging
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format='%(levelname)s: %(message)s',
        force=True
    )

    logging.info(f"Starting end-to-end generation process...")
    logging.info(f"Input: {args.input_file}")
    logging.info(f"Output: {args.output_file}")
    logging.info(f"Manufacturer: {args.manufacturer} | Model: {args.model}")

    if not os.path.exists(args.input_file):
        logging.error(f"Input file not found: {args.input_file}")
        sys.exit(1)

    ext = os.path.splitext(args.input_file)[1].lower()

    # Log warnings about options that don't match input file type
    if args.pages and ext != '.pdf':
        logging.warning("--pages option is only applicable for PDF files. Ignoring.")
    if args.sheet and ext not in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        logging.warning("--sheet option is only applicable for Excel files. Ignoring.")

    # Load custom mapping if provided
    mapping = {}
    if args.mapping:
        try:
            with open(args.mapping, 'r', encoding='utf-8') as f:
                mapping = json.load(f)
            logging.info(f"Loaded custom mapping from {args.mapping}")
        except (OSError, ValueError) as e:
            logging.error(f"Error reading mapping file: {e}")
            sys.exit(1)

    extractor = Extractor(mapping)

    pages = None
    if args.pages and ext == '.pdf':
        try:
            pages = [int(p.strip()) for p in args.pages.split(',')]
        except ValueError:
            logging.error("Invalid format for --pages. Expected comma-separated integers.")
            sys.exit(1)

    # Perform extraction based on file extension
    logging.info("Extracting data from documentation...")
    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        raw = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf':
        raw = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv':
        raw = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml':
        raw = extractor.extract_from_xml(args.input_file)
    else:
        logging.error(f"Unsupported file extension: {ext}")
        sys.exit(1)

    has_data, raw_peeked = peek_generator(raw)
    if not has_data:
        logging.error("No data extracted from the input document.")
        sys.exit(1)

    # Map columns to internal fields and apply address offsets
    logging.info("Mapping and cleaning registers...")
    mapped = extractor.map_and_clean(raw_peeked, args.address_offset)
    first, mapped_peeker = peek_generator(mapped)
    if not first:
        logging.error("No valid registers could be extracted or mapped.")
        sys.exit(1)

    # Prepare configuration for definition file generation
    config = GeneratorConfig(
        input_file=args.input_file,
        output=args.output_file,
        manufacturer=args.manufacturer,
        model=args.model,
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        address_offset=0,  # Offset already applied during map_and_clean
        template=False
    )

    # Generate definition file
    logging.info("Generating WebdynSunPM definition CSV...")
    run_generator(config, input_data=mapped_peeker)

    # Post-generation validation
    logging.info("Performing post-generation validation on the output CSV...")
    generator = Generator()
    is_valid = generator.validate_csv(args.output_file, strict=True)
    if is_valid:
        logging.info("Validation successful: The generated definition file is fully compliant!")
    else:
        logging.error("Validation failed: The generated definition file contains errors.")
        sys.exit(1)

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
