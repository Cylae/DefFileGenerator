#!/usr/bin/env python3
"""
Robust programmatic execution script to generate and validate WebdynSunPM definition files (.csv)
using the DefFileGenerator package.
"""

import os
import sys
import argparse
import logging
import re
from DefFileGenerator.extractor import Extractor, peek_generator
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator

def setup_logging(verbose: bool = False):
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format='%(levelname)s: %(message)s',
        force=True
    )

def generate_webdyn_definition(
    input_file: str,
    output_file: str,
    manufacturer: str,
    model: str,
    protocol: str = "modbusRTU",
    category: str = "Inverter",
    address_offset: int = 0,
    strict_validation: bool = True
) -> bool:
    """
    Extracts registers from a manufacturer documentation file (PDF, Excel, CSV, or XML)
    and generates a validated WebdynSunPM definition CSV file.

    Args:
        input_file: Path to the input documentation map
        output_file: Path to save the generated WebdynSunPM definition CSV
        manufacturer: Manufacturer name (e.g., "Huawei")
        model: Model name (e.g., "SUN2000")
        protocol: Protocol string (default "modbusRTU")
        category: Category string (default "Inverter")
        address_offset: Value to shift all register addresses by (default 0)
        strict_validation: Whether to fail on warnings like address overlaps (default True)

    Returns:
        bool: True if generation and validation succeeded, False otherwise
    """
    if not os.path.exists(input_file):
        logging.error(f"Input file not found: {input_file}")
        return False

    ext = os.path.splitext(input_file)[1].lower()
    extractor = Extractor()

    logging.info(f"Step 1: Extracting raw register data from: {input_file}")
    try:
        if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
            raw_data = extractor.extract_from_excel(input_file)
        elif ext == '.pdf':
            raw_data = extractor.extract_from_pdf(input_file)
        elif ext == '.csv':
            raw_data = extractor.extract_from_csv(input_file)
        elif ext == '.xml':
            raw_data = extractor.extract_from_xml(input_file)
        else:
            logging.error(f"Unsupported file format: {ext}")
            return False
    except Exception as e:
        logging.error(f"Failed to extract registers due to error: {e}")
        return False

    has_data, raw_data_peeked = peek_generator(raw_data)
    if not has_data:
        logging.error("No register data could be extracted from the file.")
        return False

    logging.info("Step 2: Cleaning and mapping fields (addresses, types, tags)...")
    mapped_gen = extractor.map_and_clean(raw_data_peeked, address_offset)
    has_regs, mapped_peeked = peek_generator(mapped_gen)
    if not has_regs:
        logging.error("No valid registers mapped after cleaning step.")
        return False

    logging.info(f"Step 3: Writing WebdynSunPM definition file to: {output_file}")
    config = GeneratorConfig(
        input_file=input_file,
        output=output_file,
        manufacturer=manufacturer,
        model=model,
        protocol=protocol,
        category=category,
        address_offset=0  # Already applied during extraction mapping
    )

    try:
        run_generator(config, input_data=mapped_peeked)
    except Exception as e:
        logging.error(f"Error during WebdynSunPM file generation: {e}")
        return False

    logging.info("Step 4: Validating the generated definition file...")
    generator = Generator()
    is_valid = generator.validate_csv(output_file, strict=strict_validation)

    if is_valid:
        logging.info(f"Success! Definition file successfully generated and validated at '{output_file}'")
        return True
    else:
        logging.error("Validation failed! The generated file contains critical errors or overlaps.")
        return False

def main():
    parser = argparse.ArgumentParser(description="Programmatic WebdynSunPM Definition File Generator")
    parser.add_argument("input_file", help="Path to input documentation (PDF, Excel, CSV, XML)")
    parser.add_argument("output_file", help="Path to save output WebdynSunPM CSV")
    parser.add_argument("--manufacturer", required=True, help="Manufacturer name")
    parser.add_argument("--model", required=True, help="Model name")
    parser.add_argument("--protocol", default="modbusRTU", help="Protocol (default: modbusRTU)")
    parser.add_argument("--category", default="Inverter", help="Device category (default: Inverter)")
    parser.add_argument("--address-offset", type=int, default=0, help="Shift all register addresses (default: 0)")
    parser.add_argument("--no-strict", dest="strict", action="store_false", help="Disable strict validation (do not fail on address overlaps)")
    parser.add_argument("-v", "--verbose", action="store_true", help="Enable verbose/debug logging")

    args = parser.parse_args()
    setup_logging(args.verbose)

    success = generate_webdyn_definition(
        input_file=args.input_file,
        output_file=args.output_file,
        manufacturer=args.manufacturer,
        model=args.model,
        protocol=args.protocol,
        category=args.category,
        address_offset=args.address_offset,
        strict_validation=args.strict
    )
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()
