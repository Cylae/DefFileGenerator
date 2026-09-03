#!/usr/bin/env python3
"""
Primary Command Line Interface for WebdynSunPM Definition Tool.

Provides sub-commands (`run`, `extract`, `generate`, `validate`) to extract registers
from documentation files, convert them into WebdynSunPM format, and validate output files.
"""

import argparse
import csv
import json
import logging
import os
import re
import sys

# Ensure parent directory is in sys.path to support direct and packaged executions
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from DefFileGenerator.def_gen import Generator, GeneratorConfig, peek_generator, run_generator
from DefFileGenerator.extractor import Extractor

try:
    from DefFileGenerator import __version__
except ImportError:  # pragma: no cover - fallback for direct execution
    __version__ = "0.0.0"

# Documentation formats the extractor knows how to read.
EXCEL_EXTENSIONS = (".xlsx", ".xlsm", ".xltx", ".xltm")
SUPPORTED_EXTENSIONS = EXCEL_EXTENSIONS + (".pdf", ".csv", ".xml")


def setup_logging(verbose=False, quiet=False):
    """Configures root logging. ``quiet`` suppresses INFO progress messages."""
    if quiet:
        level = logging.WARNING
    elif verbose:
        level = logging.DEBUG
    else:
        level = logging.INFO
    logging.basicConfig(level=level, format="%(levelname)s: %(message)s", force=True)


def _guard_output(args):
    """Refuses to silently overwrite an existing definition file."""
    output = getattr(args, "output", None)
    if output and os.path.exists(output) and not getattr(args, "force", False):
        logging.error(
            f"Output file already exists: {output}. "
            f"Re-run with --force to overwrite it, or choose another path with -o."
        )
        sys.exit(1)


def _perform_extraction(args):
    input_file = getattr(args, "input_file", None)
    if not input_file:
        logging.error("Input file is required.")
        sys.exit(1)
    if not os.path.exists(input_file):
        logging.error(f"Input file not found: {input_file}")
        sys.exit(1)

    mapping = {}
    mapping_path = getattr(args, "mapping", None)
    if mapping_path:
        try:
            with open(mapping_path) as f:
                mapping = json.load(f)
        except (OSError, ValueError) as e:
            logging.error(f"Error reading mapping file: {e}")
            sys.exit(1)

    extractor = Extractor(mapping)
    ext = os.path.splitext(input_file)[1].lower()
    address_offset = getattr(args, "address_offset", 0)
    pages_arg = getattr(args, "pages", None)
    sheet_arg = getattr(args, "sheet", None)

    if ext in EXCEL_EXTENSIONS:
        raw_data = extractor.extract_from_excel(input_file, sheet_arg)
    elif ext == ".pdf":
        pages = None
        if pages_arg:
            try:
                pages = [int(p.strip()) for p in pages_arg.split(",")]
            except ValueError:
                logging.error("Invalid format for --pages. Expected comma-separated integers.")
                sys.exit(1)
        raw_data = extractor.extract_from_pdf(input_file, pages)
    elif ext == ".csv":
        raw_data = extractor.extract_from_csv(input_file)
    elif ext == ".xml":
        raw_data = extractor.extract_from_xml(input_file)
    else:
        logging.error(
            f"Unsupported file type '{ext or '(none)'}'. "
            f"Supported formats: {', '.join(SUPPORTED_EXTENSIONS)}."
        )
        sys.exit(1)

    has_data, raw_data_peeked = peek_generator(raw_data)
    if not has_data:
        logging.error(
            f"No tabular data found in {input_file}. "
            f"For Excel try --sheet, for PDF try --pages to target the register table."
        )
        sys.exit(1)

    mapped_gen = extractor.map_and_clean(raw_data_peeked, address_offset)
    has_regs, mapped_peeked = peek_generator(mapped_gen)
    if not has_regs:
        logging.error(
            "Tables were found but no register columns could be identified. "
            "Supply a column map with --mapping (see README) or check that the "
            "sheet contains Address and Name headers."
        )
        sys.exit(1)

    return list(mapped_peeked)


def extract_command(args):
    mapped_data = _perform_extraction(args)
    first, mapped_data_iter = peek_generator(mapped_data)
    if not first:
        logging.error("No registers extracted.")
        sys.exit(1)

    output = getattr(args, "output", None)
    fieldnames = [
        "Name",
        "Tag",
        "RegisterType",
        "Address",
        "Type",
        "Factor",
        "Offset",
        "Unit",
        "Action",
        "ScaleFactor",
    ]

    if output:
        f = open(output, "w", newline="", encoding="utf-8")
    else:
        f = sys.stdout

    try:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(mapped_data_iter)
    finally:
        if output:
            f.close()
            logging.info(f"Extraction complete. Saved to {output}")


def validate_command(args):
    """Validates a definition file; ``--lenient`` downgrades overlaps to warnings."""
    if not os.path.exists(args.input_file):
        logging.error(f"File not found: {args.input_file}")
        sys.exit(1)
    generator = Generator()
    strict = not getattr(args, "lenient", False)
    if generator.validate_csv(args.input_file, strict=strict):
        logging.info(f"Validation successful: {args.input_file}")
    else:
        logging.error(
            f"Validation failed: {args.input_file}. "
            f"Review the errors above; re-run with --lenient to treat "
            f"address overlaps as warnings."
        )
        sys.exit(1)


def generate_command(args):
    template = getattr(args, "template", False)
    template_mode = getattr(args, "template_mode", "input")

    config = GeneratorConfig(
        input_file=getattr(args, "input_file", None),
        output=getattr(args, "output", None),
        manufacturer=getattr(args, "manufacturer", "Manufacturer"),
        model=getattr(args, "model", "Model"),
        protocol=getattr(args, "protocol", "modbusRTU"),
        category=getattr(args, "category", "Inverter"),
        forced_write=getattr(args, "forced_write", ""),
        address_offset=getattr(args, "address_offset", 0),
        template=template,
        template_mode=template_mode,
    )
    run_generator(config)


def run_command(args):
    template = getattr(args, "template", False)
    mapped_data = None
    if not template:
        mapped_data = _perform_extraction(args)
        first, mapped_data_iter = peek_generator(mapped_data)
        if not first:
            logging.error("No registers extracted.")
            sys.exit(1)
        mapped_data = mapped_data_iter

    m_name = getattr(args, "manufacturer", None)
    m_model = getattr(args, "model", None)
    if not template and (not m_name or not m_model):
        logging.error("--manufacturer and --model are required for run mode.")
        sys.exit(1)

    m_name = m_name or "Manufacturer"
    m_model = m_model or "Model"
    output_file = getattr(args, "output", None)
    if not output_file and not template:
        # Sanitize name and model for default output filename
        sanitized_mfg = re.sub(r"[^a-zA-Z0-9]", "_", m_name).lower()
        sanitized_model = re.sub(r"[^a-zA-Z0-9]", "_", m_model).lower()
        output_file = f"{sanitized_mfg}_{sanitized_model}_definition.csv"

    config = GeneratorConfig(
        input_file=getattr(args, "input_file", None),
        output=output_file,
        manufacturer=m_name,
        model=m_model,
        protocol=getattr(args, "protocol", "modbusRTU"),
        category=getattr(args, "category", "Inverter"),
        forced_write=getattr(args, "forced_write", ""),
        address_offset=0,  # Already applied during extraction
        template=template,
        template_mode=getattr(args, "template_mode", "input"),
    )
    run_generator(config, input_data=mapped_data)

    if template or not output_file or not os.path.exists(output_file):
        return

    logging.info(f"Definition written to {output_file}")

    # Self-check the artefact we just produced so the operator learns about
    # overlaps or malformed rows now rather than when the file is imported
    # into a device.
    if getattr(args, "no_validate", False):
        return
    if Generator().validate_csv(output_file, strict=True):
        logging.info("Post-generation validation passed.")
    else:
        logging.warning(
            "Generated file has validation warnings (see above). "
            "Inspect it before importing, or re-check with "
            f"'deffilegen validate {output_file}'."
        )


def _run_cli(argv=None):
    if argv is None:
        argv = sys.argv[1:]
    else:
        # Strip script name if present as first element
        if argv and (
            argv[0].endswith("main.py")
            or argv[0].endswith("doc_to_webdyn.py")
            or argv[0] == "main.py"
            or argv[0] == "doc_to_webdyn.py"
        ):
            argv = argv[1:]

    parser = argparse.ArgumentParser(
        prog="deffilegen",
        description=(
            "WebdynSunPM definition tool: extract Modbus register maps from "
            "manufacturer documentation (PDF, Excel, CSV, XML) and emit "
            "validated WebdynSunPM definition files."
        ),
        epilog=(
            "Examples:\n"
            "  deffilegen run datasheet.pdf --manufacturer Huawei --model SUN2000-50KTL\n"
            '  deffilegen extract book.xlsx --sheet "Holding Registers" -o registers.csv\n'
            "  deffilegen generate registers.csv --manufacturer SMA --model STP-5000TL -o sma.csv\n"
            "  deffilegen validate sma.csv\n"
            "  deffilegen generate --template -o starter.csv\n"
            "\nExit codes: 0 success, 1 error, 2 bad usage, 130 interrupted."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="Enable debug-level logging.")
    parser.add_argument(
        "-q", "--quiet", action="store_true", help="Only report warnings and errors."
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")

    # Shared verbosity flags, attached to every sub-parser as well, so both
    # "deffilegen -v run ..." and "deffilegen run ... -v" work. Operators
    # routinely append -v after the command; failing there is a papercut.
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("-v", "--verbose", action="store_true", help="Enable debug-level logging.")
    common.add_argument(
        "-q", "--quiet", action="store_true", help="Only report warnings and errors."
    )

    subparsers = parser.add_subparsers(
        dest="command", metavar="COMMAND", help="Sub-command to run."
    )

    # Validate
    parser_validate = subparsers.add_parser(
        "validate",
        parents=[common],
        help="Check an existing definition file for errors.",
        description="Validates tag uniqueness, type codes, address ranges and register overlaps.",
    )
    parser_validate.add_argument(
        "input_file", metavar="DEFINITION_CSV", help="WebdynSunPM definition CSV to validate."
    )
    parser_validate.add_argument(
        "--lenient",
        action="store_true",
        help="Report address overlaps as warnings instead of failing.",
    )

    # Extract
    parser_extract = subparsers.add_parser(
        "extract",
        parents=[common],
        help="Extract a register map to an intermediate CSV.",
        description=(
            "Reads documentation and writes the detected registers without WebdynSunPM formatting."
        ),
    )
    parser_extract.add_argument(
        "input_file", metavar="SOURCE", help="Documentation file (.pdf, .xlsx, .xlsm, .csv, .xml)."
    )
    parser_extract.add_argument(
        "-o", "--output", metavar="FILE", help="Destination CSV (default: standard output)."
    )
    parser_extract.add_argument(
        "--mapping", metavar="JSON", help="JSON file overriding automatic column detection."
    )
    parser_extract.add_argument(
        "--sheet", metavar="NAME", help="Excel worksheet to read (default: every sheet)."
    )
    parser_extract.add_argument(
        "--pages", metavar="LIST", help="Comma-separated PDF page numbers, e.g. 12,13,14."
    )
    parser_extract.add_argument(
        "--address-offset",
        type=int,
        default=0,
        metavar="N",
        help="Shift every extracted address by N (default: 0).",
    )
    parser_extract.add_argument(
        "--force", action="store_true", help="Overwrite the output file if it already exists."
    )

    # Generate
    parser_generate = subparsers.add_parser(
        "generate",
        parents=[common],
        help="Build a definition file from an intermediate CSV.",
        description="Converts an intermediate register CSV into a WebdynSunPM definition file.",
    )
    parser_generate.add_argument(
        "input_file",
        nargs="?",
        metavar="REGISTERS_CSV",
        help='Intermediate CSV produced by "extract".',
    )
    parser_generate.add_argument(
        "--manufacturer",
        metavar="NAME",
        help="Manufacturer recorded in the header (required unless --template).",
    )
    parser_generate.add_argument(
        "--model", metavar="NAME", help="Model recorded in the header (required unless --template)."
    )
    parser_generate.add_argument(
        "-o", "--output", metavar="FILE", help="Destination CSV (default: standard output)."
    )
    parser_generate.add_argument(
        "--template",
        action="store_true",
        help="Emit a starter template instead of converting a file.",
    )
    parser_generate.add_argument(
        "--template-mode",
        choices=["input", "definition"],
        default="input",
        help="Template flavour to emit (default: input).",
    )
    parser_generate.add_argument(
        "--protocol",
        default="modbusRTU",
        metavar="NAME",
        help="Protocol header field (default: modbusRTU).",
    )
    parser_generate.add_argument(
        "--category",
        default="Inverter",
        metavar="NAME",
        help="Device category header field (default: Inverter).",
    )
    parser_generate.add_argument(
        "--forced-write", default="", metavar="VALUE", help="Optional forced-write header field."
    )
    parser_generate.add_argument(
        "--address-offset",
        type=int,
        default=0,
        metavar="N",
        help="Shift every address by N (default: 0).",
    )
    parser_generate.add_argument(
        "--force", action="store_true", help="Overwrite the output file if it already exists."
    )

    # Run
    parser_run = subparsers.add_parser(
        "run",
        parents=[common],
        help="Extract and generate in a single step.",
        description="End-to-end conversion from manufacturer documentation to a definition file.",
    )
    parser_run.add_argument(
        "input_file",
        nargs="?",
        metavar="SOURCE",
        help="Documentation file (.pdf, .xlsx, .xlsm, .csv, .xml).",
    )
    parser_run.add_argument(
        "--manufacturer",
        metavar="NAME",
        help="Manufacturer recorded in the header (required unless --template).",
    )
    parser_run.add_argument(
        "--model", metavar="NAME", help="Model recorded in the header (required unless --template)."
    )
    parser_run.add_argument(
        "-o",
        "--output",
        metavar="FILE",
        help="Destination CSV (default: <manufacturer>_<model>_definition.csv).",
    )
    parser_run.add_argument(
        "--template",
        action="store_true",
        help="Emit a starter template instead of converting a file.",
    )
    parser_run.add_argument(
        "--template-mode",
        choices=["input", "definition"],
        default="input",
        help="Template flavour to emit (default: input).",
    )
    parser_run.add_argument(
        "--mapping", metavar="JSON", help="JSON file overriding automatic column detection."
    )
    parser_run.add_argument(
        "--sheet", metavar="NAME", help="Excel worksheet to read (default: every sheet)."
    )
    parser_run.add_argument(
        "--pages", metavar="LIST", help="Comma-separated PDF page numbers, e.g. 12,13,14."
    )
    parser_run.add_argument(
        "--protocol",
        default="modbusRTU",
        metavar="NAME",
        help="Protocol header field (default: modbusRTU).",
    )
    parser_run.add_argument(
        "--category",
        default="Inverter",
        metavar="NAME",
        help="Device category header field (default: Inverter).",
    )
    parser_run.add_argument(
        "--forced-write", default="", metavar="VALUE", help="Optional forced-write header field."
    )
    parser_run.add_argument(
        "--address-offset",
        type=int,
        default=0,
        metavar="N",
        help="Shift every address by N (default: 0).",
    )
    parser_run.add_argument(
        "--force", action="store_true", help="Overwrite the output file if it already exists."
    )
    parser_run.add_argument(
        "--no-validate", action="store_true", help="Skip the post-generation validation pass."
    )

    args = parser.parse_args(argv)

    if not args.command:
        parser.print_help(sys.stderr)
        sys.exit(2)

    setup_logging(args.verbose, getattr(args, "quiet", False))

    # Validate pages/sheet parameters depending on input file type
    input_file = getattr(args, "input_file", None)
    if input_file:
        ext = os.path.splitext(input_file)[1].lower()
        if getattr(args, "pages", None) and ext != ".pdf":
            logging.warning("--pages is only applicable for PDF files. Ignoring.")
        if getattr(args, "sheet", None) and ext not in EXCEL_EXTENSIONS:
            logging.warning("--sheet is only applicable for Excel files. Ignoring.")

    # Validation checks for required parameters
    if args.command in ["generate", "run"] and not getattr(args, "template", False):
        missing = [
            flag
            for flag, value in (
                ("--manufacturer", getattr(args, "manufacturer", None)),
                ("--model", getattr(args, "model", None)),
            )
            if not value
        ]
        if missing:
            logging.error(
                f"'{args.command}' requires {' and '.join(missing)}. "
                f"Example: deffilegen {args.command} INPUT "
                f"--manufacturer Huawei --model SUN2000-50KTL"
            )
            sys.exit(1)
        if not getattr(args, "input_file", None):
            logging.error(
                f"'{args.command}' requires an input file. "
                f"Run 'deffilegen {args.command} --help' for usage."
            )
            sys.exit(1)
        if not os.path.exists(args.input_file):
            logging.error(f"Input file not found: {args.input_file}")
            sys.exit(1)
        _guard_output(args)

    if args.command == "extract":
        if not os.path.exists(args.input_file):
            logging.error(f"Input file not found: {args.input_file}")
            sys.exit(1)
        _guard_output(args)
        extract_command(args)
    elif args.command == "validate":
        validate_command(args)
    elif args.command == "generate":
        generate_command(args)
    elif args.command == "run":
        run_command(args)


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
