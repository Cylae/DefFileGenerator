#!/usr/bin/env python3
"""
Robust programmatic execution script to generate and validate WebdynSunPM definition files (.csv)
using the DefFileGenerator package.
"""

import logging
import os
import sys
from dataclasses import dataclass
from typing import Optional, Union

from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator
from DefFileGenerator.extractor import Extractor, peek_generator


@dataclass
class WebdynDefConfig:
    input_file: str
    output_file: str
    manufacturer: str
    model: str
    protocol: str = "modbusRTU"
    category: str = "Inverter"
    address_offset: int = 0
    strict_validation: bool = True


def setup_logging():
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s", force=True)


def generate_webdyn_definition(
    input_file: Union[str, WebdynDefConfig],
    output_file: Optional[str] = None,
    manufacturer: Optional[str] = None,
    model: Optional[str] = None,
    protocol: str = "modbusRTU",
    category: str = "Inverter",
    address_offset: int = 0,
    strict_validation: bool = True,
) -> bool:
    """
    Extracts registers from a manufacturer documentation file (PDF, Excel, CSV, or XML)
    and generates a validated WebdynSunPM definition CSV file.

    Args:
        input_file: Path to the input documentation map OR a WebdynDefConfig instance
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
    if isinstance(input_file, WebdynDefConfig):
        cfg = input_file
        input_file_str = cfg.input_file
        output_file_str = cfg.output_file
        mfg = cfg.manufacturer
        mdl = cfg.model
        proto = cfg.protocol
        cat = cfg.category
        offset = cfg.address_offset
        strict = cfg.strict_validation
    else:
        input_file_str = str(input_file)
        output_file_str = str(output_file) if output_file is not None else ""
        mfg = str(manufacturer) if manufacturer is not None else ""
        mdl = str(model) if model is not None else ""
        proto = protocol
        cat = category
        offset = address_offset
        strict = strict_validation

    if not os.path.exists(input_file_str):
        logging.error(f"Input file not found: {input_file_str}")
        return False

    ext = os.path.splitext(input_file_str)[1].lower()
    extractor = Extractor()

    logging.info(f"Step 1: Extracting raw register data from: {input_file_str}")
    try:
        if ext in [".xlsx", ".xlsm", ".xltx", ".xltm"]:
            raw_data = extractor.extract_from_excel(input_file_str)
        elif ext == ".pdf":
            raw_data = extractor.extract_from_pdf(input_file_str)
        elif ext == ".csv":
            raw_data = extractor.extract_from_csv(input_file_str)
        elif ext == ".xml":
            raw_data = extractor.extract_from_xml(input_file_str)
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
    mapped_gen = extractor.map_and_clean(raw_data_peeked, offset)
    has_regs, mapped_peeked = peek_generator(mapped_gen)
    if not has_regs:
        logging.error("No valid registers mapped after cleaning step.")
        return False

    logging.info(f"Step 3: Writing WebdynSunPM definition file to: {output_file_str}")
    config = GeneratorConfig(
        input_file=input_file_str,
        output=output_file_str,
        manufacturer=mfg,
        model=mdl,
        protocol=proto,
        category=cat,
        address_offset=0,  # Already applied during extraction mapping
    )

    try:
        run_generator(config, input_data=mapped_peeked)
    except Exception as e:
        logging.error(f"Error during WebdynSunPM file generation: {e}")
        return False

    logging.info("Step 4: Validating the generated definition file...")
    generator = Generator()
    is_valid = generator.validate_csv(output_file_str, strict=strict)

    if is_valid:
        logging.info(
            f"Success! Definition file successfully generated and validated at '{output_file_str}'"
        )
        return True
    else:
        logging.error("Validation failed! The generated file contains critical errors or overlaps.")
        return False


def main():
    setup_logging()

    if len(sys.argv) > 1 and sys.argv[1] in ("-h", "--help"):
        print("Usage:")
        print(
            "  python3 generate_webdyn_def.py <input_file> <output_file> <manufacturer> <model> [protocol] [category]"
        )
        print("  python3 generate_webdyn_def.py (runs demo mode if no arguments provided)")
        sys.exit(0)

    # We can run a demo if no arguments are passed, or print usage
    if len(sys.argv) < 5:
        print("Usage:")
        print(
            "  python3 generate_webdyn_def.py <input_file> <output_file> <manufacturer> <model> [options]"
        )
        print("\nDemo execution using sample registers:")

        sample_in = "sample_register_map.csv"
        sample_out = "sample_output_definition.csv"

        # Create a sample register CSV if not present, to run the demo
        if not os.path.exists(sample_in):
            with open(sample_in, "w", encoding="utf-8") as f:
                f.write("Register,Name,Data Type,Unit,Scale,Access\n")
                f.write("40001,AC Power,uint16,W,1,R\n")
                f.write("40002,DC Voltage,uint16,V,0.1,R\n")
                f.write("40003,Temperature,int16,°C,0.1,R\n")

        logging.info("Running demo generation...")
        success = generate_webdyn_definition(
            input_file=sample_in, output_file=sample_out, manufacturer="DemoMfg", model="DemoModel"
        )
        sys.exit(0 if success else 1)
        return

    input_file = sys.argv[1]
    output_file = sys.argv[2]
    manufacturer = sys.argv[3]
    model = sys.argv[4]

    # Optional arguments
    protocol = sys.argv[5] if len(sys.argv) > 5 else "modbusRTU"
    category = sys.argv[6] if len(sys.argv) > 6 else "Inverter"

    success = generate_webdyn_definition(
        input_file=input_file,
        output_file=output_file,
        manufacturer=manufacturer,
        model=model,
        protocol=protocol,
        category=category,
    )
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
