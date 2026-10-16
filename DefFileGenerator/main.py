#!/usr/bin/env python3
"""
Primary Command Line Interface for WebdynSunPM Definition Tool.

Provides unified sub-commands:
    - run: End-to-end extraction from documentation and Webdyn CSV generation.
    - extract: Parses input files into a clean intermediate tabular register CSV.
    - generate: Converts an intermediate register CSV into a validated Webdyn definition CSV.
    - validate: Audits an existing Webdyn definition file for compliance and syntax errors.
    - gui: Launches the native Windows 11 desktop application.
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import os
import re
import sys
from collections.abc import Iterator
from contextlib import nullcontext
from pathlib import Path
from typing import Any

# Ensure parent directory is in sys.path to support direct and packaged executions
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from DefFileGenerator.def_gen import (
    Generator,
    GeneratorConfig,
    peek_generator,
    run_generator,
)
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.io_utils import paths_refer_to_same_file, staged_text_output

ALLOWED_EXTENSIONS = frozenset({".pdf", ".xlsx", ".xlsm", ".xltx", ".xltm", ".csv", ".xml"})
RE_SLUGIFY = re.compile(r"[^a-zA-Z0-9]")


def _guard_source_output(input_file: str | None, output: str | None) -> None:
    if input_file and output and paths_refer_to_same_file(input_file, output):
        logging.error("The output file must differ from the source file, even with --force.")
        sys.exit(1)


def setup_logging(verbose: bool = False, quiet: bool = False) -> None:
    """
    Configures console logging severity.

    Args:
        verbose: Enables DEBUG level logging.
        quiet: Suppresses INFO logging; only WARNING and ERROR are displayed.
    """
    level = logging.DEBUG if verbose else (logging.WARNING if quiet else logging.INFO)
    logging.basicConfig(level=level, format="%(levelname)s: %(message)s", force=True)


def _get_default_output(manufacturer: str, model: str) -> str:
    """
    Generates a standardized filename from manufacturer and model strings.

    Args:
        manufacturer: Manufacturer brand name (e.g. "Huawei").
        model: Equipment model identifier (e.g. "SUN2000").

    Returns:
        str: Output definition filename formatted as `<mfg>_<model>_definition.csv`.
    """
    m = RE_SLUGIFY.sub("_", manufacturer).lower()
    md = RE_SLUGIFY.sub("_", model).lower()
    return f"{m}_{md}_definition.csv"


def _perform_extraction(args: argparse.Namespace) -> Iterator[dict[str, Any]]:
    """
    Validates input arguments and executes multi-format register extraction.

    Args:
        args: Parsed command line arguments.

    Returns:
        Iterator[dict[str, Any]]: Stream of extracted and normalized register dictionaries.

    Raises:
        SystemExit: If input file is missing, unsupported, or contains no registers.
    """
    input_file = getattr(args, "input_file", None)
    if not input_file:
        logging.error("Input file is required.")
        sys.exit(1)
    if not os.path.exists(input_file):
        logging.error(f"Input file not found: {input_file}")
        sys.exit(1)

    ext = os.path.splitext(input_file)[1].lower()
    if ext not in ALLOWED_EXTENSIONS:
        logging.error(
            f"Unsupported file type: {ext}. Supported formats are: .pdf, .xlsx, .xlsm, .xltx, .xltm, .csv, .xml"
        )
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
    address_offset = getattr(args, "address_offset", 0)
    pages = getattr(args, "pages", None)
    sheet = getattr(args, "sheet", None)

    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        raw_data = extractor.extract_from_excel(input_file, sheet)
    elif ext == ".pdf":
        raw_data = extractor.extract_from_pdf(input_file, pages)
    elif ext == ".csv":
        raw_data = extractor.extract_from_csv(input_file)
    elif ext == ".xml":
        raw_data = extractor.extract_from_xml(input_file)
    else:
        logging.error(f"Unsupported file type: {ext}")
        sys.exit(1)

    has_data, raw_peeked = peek_generator(raw_data)
    if not has_data:
        logging.error("No data extracted.")
        sys.exit(1)

    mapped_gen = extractor.map_and_clean(raw_peeked, address_offset)
    has_regs, mapped_peeked = peek_generator(mapped_gen)
    if not has_regs:
        logging.error("No registers extracted.")
        sys.exit(1)

    return mapped_peeked


def extract_command(args: argparse.Namespace) -> None:
    """
    Subcommand handler for 'extract': parses documentation and outputs an intermediate CSV.

    Args:
        args: Parsed command line arguments.
    """
    output = getattr(args, "output", None)
    _guard_source_output(getattr(args, "input_file", None), output)
    if output and os.path.exists(output) and not getattr(args, "force", False):
        logging.error(f"Output file '{output}' already exists. Use --force to overwrite.")
        sys.exit(1)

    mapped_data = _perform_extraction(args)
    first, mapped_data_iter = peek_generator(mapped_data)
    if not first:
        logging.error("No registers extracted.")
        sys.exit(1)

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

    output_context = staged_text_output(output) if output else nullcontext(sys.stdout)
    with output_context as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(
            {field: Generator.sanitize_csv_field(row.get(field, "")) for field in fieldnames}
            for row in mapped_data_iter
        )
    if output:
        logging.info(f"Extraction complete. Saved to {output}")


def validate_command(args: argparse.Namespace) -> None:
    """
    Subcommand handler for 'validate': audits a Webdyn definition CSV file.

    Args:
        args: Parsed command line arguments.
    """
    input_file = getattr(args, "input_file", None)
    if not input_file or not os.path.exists(input_file):
        logging.error(f"File not found: {input_file}")
        sys.exit(1)

    generator = Generator()
    strict = not getattr(args, "lenient", False)
    report = generator.validate_csv_detailed(input_file, strict=strict)
    if report.is_valid:
        logging.info(f"Validation successful: {input_file}")
    else:
        logging.error(f"Validation failed: {input_file}")
        sys.exit(1)


def tag_command(args: argparse.Namespace) -> None:
    """
    Subcommand handler for 'tag': enriches a definition CSV with canonical SunPM tags.

    Args:
        args: Parsed command line arguments.
    """
    input_file = getattr(args, "input_file", None)
    if not input_file or not os.path.exists(input_file):
        logging.error(f"File not found: {input_file}")
        sys.exit(1)

    output = getattr(args, "output", None) or input_file
    equipment_type = getattr(args, "category", "")
    overwrite = getattr(args, "overwrite", False)
    normalize_legacy = not getattr(args, "no_normalize", False)

    from DefFileGenerator.knowledge_miner import enrich_definition_file

    count = enrich_definition_file(
        input_path=input_file,
        output_path=output,
        equipment_type=equipment_type,
        overwrite_existing=overwrite,
        normalize_legacy=normalize_legacy,
    )
    logging.info(
        f"SunPM Tag enrichment complete for {input_file}. Updated {count} tag(s). Saved to {output}"
    )


def audit_command(args: argparse.Namespace) -> None:
    """
    Subcommand handler for 'audit': verifies syntax, address collisions, tags, and units.
    Optionally applies automated non-destructive fixes.

    Args:
        args: Parsed command line arguments.
    """
    from dataclasses import asdict

    from DefFileGenerator.knowledge_miner import (
        audit_definition_file,
        audit_definitions_directory,
        auto_fix_definition_file,
    )

    target = getattr(args, "target", "Definitions")
    strict = getattr(args, "strict", False)
    strict_overlap = getattr(args, "strict_overlap", False)
    do_fix = getattr(args, "fix", False)
    as_json = getattr(args, "json", False)
    output = getattr(args, "output", None)

    if not os.path.exists(target):
        logging.error(f"Target path not found: {target}")
        sys.exit(1)

    if os.path.isfile(target):
        fix_stats = None
        if do_fix:
            fix_target = output or target
            fix_stats = auto_fix_definition_file(target, output_path=fix_target)
            logging.info(
                f"Auto-fix applied to {target}: "
                f"{fix_stats['tags_normalized']} normalized, "
                f"{fix_stats['tags_inferred']} inferred, "
                f"{fix_stats['duplicates_removed']} duplicates removed."
            )
            report_target = fix_target
        else:
            report_target = target

        report = audit_definition_file(report_target, strict_overlap=strict_overlap)

        if as_json:
            out_data: dict[str, Any] = asdict(report)
            if fix_stats:
                out_data["fix_applied"] = fix_stats
            print(json.dumps(out_data, indent=2))
        else:
            logging.info(f"Audited {report.filepath} ({report.register_count} registers)")
            if report.errors:
                logging.error(f"Found {len(report.errors)} ERROR(s):")
                for err in report.errors:
                    logging.error(f"  Line {err.line} [{err.code}]: {err.message}")
            if report.warnings:
                logging.warning(f"Found {len(report.warnings)} WARNING(s):")
                for warn in report.warnings:
                    logging.warning(f"  Line {warn.line} [{warn.code}]: {warn.message}")
            if report.infos and getattr(args, "verbose", False):
                logging.info(f"Found {len(report.infos)} INFO note(s):")
                for inf in report.infos:
                    logging.info(f"  Line {inf.line} [{inf.code}]: {inf.message}")

        has_failed = bool(report.errors) or (strict and bool(report.warnings))
        if has_failed:
            sys.exit(1)

    elif os.path.isdir(target):
        fix_summary = {"tags_normalized": 0, "tags_inferred": 0, "duplicates_removed": 0}
        if do_fix:
            for p in sorted(Path(target).glob("*.csv")):
                s = auto_fix_definition_file(p)
                for k, v in s.items():
                    fix_summary[k] += v
            logging.info(
                f"Batch auto-fix applied to {target}: "
                f"{fix_summary['tags_normalized']} normalized, "
                f"{fix_summary['tags_inferred']} inferred, "
                f"{fix_summary['duplicates_removed']} duplicates removed."
            )

        reports = audit_definitions_directory(target, strict_overlap=strict_overlap)
        total_files = len(reports)
        total_regs = sum(r.register_count for r in reports.values())
        total_errors = sum(len(r.errors) for r in reports.values())
        total_warnings = sum(len(r.warnings) for r in reports.values())
        total_infos = sum(len(r.infos) for r in reports.values())
        invalid_files = [f for f, r in reports.items() if not r.is_valid]

        if as_json:
            payload: dict[str, Any] = {
                "summary": {
                    "total_files": total_files,
                    "valid_files": total_files - len(invalid_files),
                    "invalid_files": len(invalid_files),
                    "total_registers": total_regs,
                    "total_errors": total_errors,
                    "total_warnings": total_warnings,
                    "total_infos": total_infos,
                },
                "reports": {f: asdict(r) for f, r in reports.items()},
            }
            if do_fix:
                payload["fix_applied"] = fix_summary
            print(json.dumps(payload, indent=2))
        else:
            logging.info(
                f"Audited {total_files} definition files ({total_regs:,} registers) in '{target}'."
            )
            logging.info(
                f"Results: {total_files - len(invalid_files)}/{total_files} valid, "
                f"{total_errors} error(s), {total_warnings} warning(s), {total_infos} info(s)."
            )
            if invalid_files:
                logging.error(f"Files with errors ({len(invalid_files)}):")
                for f in invalid_files:
                    logging.error(f"  - {f}: {len(reports[f].errors)} error(s)")
            if strict and total_warnings > 0:
                logging.error(f"--strict enabled and {total_warnings} warning(s) found.")

        has_failed = total_errors > 0 or (strict and total_warnings > 0)
        if has_failed:
            sys.exit(1)


def train_command(args: argparse.Namespace) -> None:
    """
    Subcommand handler for 'train': learns empirical metrics, data types, units,
    action codes, and official SunPM tags from definition CSVs.

    Args:
        args: Parsed command line arguments.
    """
    from DefFileGenerator.knowledge_miner import train_knowledge_from_definitions

    definitions_dir = getattr(args, "dir", "Definitions")
    export_json = getattr(args, "export_json", None)
    summary = getattr(args, "summary", False)

    if not os.path.exists(definitions_dir):
        logging.error(f"Definitions directory not found: {definitions_dir}")
        sys.exit(1)

    knowledge = train_knowledge_from_definitions(directories=definitions_dir, force_retrain=True)

    logging.info(
        f"Knowledge training successful: "
        f"{knowledge.total_definitions_mined} files, "
        f"{knowledge.total_registers_mined:,} registers mined, "
        f"{len(knowledge.name_to_tag):,} distinct variable-to-tag mappings learned."
    )

    if summary:
        top_tags = sorted(knowledge.tag_counts.items(), key=lambda x: x[1], reverse=True)[:15]
        logging.info("Top SunPM Tags in dataset:")
        for tag, count in top_tags:
            logging.info(f"  {tag:<15}: {count} occurrences")

    if export_json:
        data = knowledge.to_dict()
        with staged_text_output(export_json) as f:
            json.dump(data, f, indent=2)
        logging.info(f"Exported learned dataset knowledge to {export_json}")


def generate_command(args: argparse.Namespace) -> None:
    """
    Subcommand handler for 'generate': converts an intermediate register CSV into Webdyn format.

    Args:
        args: Parsed command line arguments.
    """
    template = getattr(args, "template", False)
    template_mode = getattr(args, "template_mode", "input")
    m_name = getattr(args, "manufacturer", None)
    m_model = getattr(args, "model", None)
    input_file = getattr(args, "input_file", None)
    output = getattr(args, "output", None)
    _guard_source_output(input_file, output)

    if not template:
        if not m_name or not m_model:
            logging.error("--manufacturer and --model are required for generate command.")
            sys.exit(1)
        if not input_file:
            logging.error("input_file is required for generate command.")
            sys.exit(1)
        if not os.path.exists(input_file):
            logging.error(f"Input file not found: {input_file}")
            sys.exit(1)

    if output and os.path.exists(output) and not getattr(args, "force", False):
        logging.error(f"Output file '{output}' already exists. Use --force to overwrite.")
        sys.exit(1)

    config = GeneratorConfig(
        input_file=input_file,
        output=output,
        manufacturer=m_name or "Manufacturer",
        model=m_model or "Model",
        protocol=getattr(args, "protocol", "modbusRTU"),
        category=getattr(args, "category", "Inverter"),
        forced_write=getattr(args, "forced_write", ""),
        address_offset=getattr(args, "address_offset", 0),
        template=template,
        template_mode=template_mode,
        strict_validation=True,
    )
    if run_generator(config) is False:
        logging.error("Definition generation failed.")
        sys.exit(1)

    report_path = getattr(args, "report", None)
    if report_path and output:
        from DefFileGenerator.report_generator import generate_single_definition_report

        generate_single_definition_report(output, output_report_path=report_path)
        logging.info(f"Technical report generated: {report_path}")


def run_command(args: argparse.Namespace) -> None:
    """
    Subcommand handler for 'run': executes end-to-end extraction, generation, and validation.

    Args:
        args: Parsed command line arguments.
    """
    template = getattr(args, "template", False)
    m_name = getattr(args, "manufacturer", None)
    m_model = getattr(args, "model", None)
    input_file = getattr(args, "input_file", None)

    if not template:
        if not m_name or not m_model:
            logging.error("--manufacturer and --model are required for run command.")
            sys.exit(1)
        if not input_file:
            logging.error("input_file is required for run command.")
            sys.exit(1)
        if not os.path.exists(input_file):
            logging.error(f"Input file not found: {input_file}")
            sys.exit(1)

    mapped_data = None
    if not template:
        mapped_data = _perform_extraction(args)
        has_regs, mapped_data = peek_generator(mapped_data)
        if not has_regs:
            logging.error("No registers extracted.")
            sys.exit(1)

    output_file = getattr(args, "output", None)
    if not output_file and not template:
        output_file = _get_default_output(str(m_name or "Manufacturer"), str(m_model or "Model"))

    _guard_source_output(input_file, output_file)

    if output_file and os.path.exists(output_file) and not getattr(args, "force", False):
        logging.error(f"Output file '{output_file}' already exists. Use --force to overwrite.")
        sys.exit(1)

    config = GeneratorConfig(
        input_file=input_file,
        output=output_file,
        manufacturer=m_name or "Manufacturer",
        model=m_model or "Model",
        protocol=getattr(args, "protocol", "modbusRTU"),
        category=getattr(args, "category", "Inverter"),
        forced_write=getattr(args, "forced_write", ""),
        address_offset=0,  # Already applied during extraction
        template=template,
        template_mode=getattr(args, "template_mode", "input"),
        strict_validation=not getattr(args, "no_validate", False),
    )
    if run_generator(config, input_data=mapped_data) is False:
        logging.error("Definition generation failed.")
        sys.exit(1)

    if output_file and not getattr(args, "no_validate", False) and not template:
        generator = Generator()
        if generator.validate_csv(output_file, strict=True):
            logging.info("Post-generation validation passed.")
            report_path = getattr(args, "report", None)
            if report_path:
                from DefFileGenerator.report_generator import generate_single_definition_report

                generate_single_definition_report(output_file, output_report_path=report_path)
                logging.info(f"Technical report generated: {report_path}")
        else:
            logging.error("Post-generation validation failed.")
            sys.exit(1)


def report_command(args: argparse.Namespace) -> None:
    """
    Subcommand handler for 'report': generates comprehensive technical reports in Markdown/JSON.

    Args:
        args: Parsed command line arguments.
    """
    from DefFileGenerator.report_generator import (
        analyze_definitions_directory,
        export_report_json,
        generate_directory_technical_report,
        generate_single_definition_report,
    )

    target = getattr(args, "target", "tagged_definitions")
    output = getattr(args, "output", None)
    json_path = getattr(args, "json", None)

    if not os.path.exists(target):
        logging.error(f"Target path not found: {target}")
        sys.exit(1)

    if os.path.isfile(target):
        out_report = output or f"{os.path.splitext(target)[0]}_report.md"
        generate_single_definition_report(target, output_report_path=out_report)
        logging.info(f"Technical report generated: {out_report}")
    elif os.path.isdir(target):
        out_report = output or os.path.join(target, "RAPPORT_TECHNIQUE_GENERATION.md")
        generate_directory_technical_report(target, output_report_path=out_report)
        logging.info(f"Directory technical report generated: {out_report}")
        if json_path:
            stats = analyze_definitions_directory(target)
            export_report_json(stats, json_path)
            logging.info(f"Structured metrics exported: {json_path}")


def _run_cli(argv: list[str] | None = None) -> None:
    """
    Parses CLI arguments and dispatches to the corresponding command handler.

    Args:
        argv: List of command line arguments (defaults to sys.argv[1:]).
    """
    if argv is None:
        argv = sys.argv[1:]
    else:
        if argv and (
            argv[0].endswith("main.py")
            or argv[0].endswith("doc_to_webdyn.py")
            or argv[0] == "main.py"
            or argv[0] == "doc_to_webdyn.py"
        ):
            argv = argv[1:]

    parser = argparse.ArgumentParser(
        prog="deffilegen",
        description="WebdynSunPM Definition Tool",
        epilog=(
            "Examples:\n"
            "  deffilegen run input.xlsx --manufacturer Huawei --model SUN2000 -o huawei.csv\n"
            "  deffilegen extract doc.pdf -o registers.csv\n"
            "  deffilegen generate registers.csv --manufacturer SMA --model STP5000 -o sma.csv\n"
            "  deffilegen validate definition.csv\n\n"
            "Exit codes:\n"
            "  0: Success\n"
            "  1: Execution error\n"
            "  2: Usage / argument error"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--version", action="version", version="deffilegen 0.2.1")
    parser.add_argument("--gui", action="store_true", help="Launch Windows 11 Desktop Application")
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose logging")
    parser.add_argument("-q", "--quiet", action="store_true", help="Quiet logging")

    subparsers = parser.add_subparsers(dest="command", help="Sub-commands")

    def add_common_flags(p: argparse.ArgumentParser) -> None:
        p.add_argument(
            "-v",
            "--verbose",
            action="store_true",
            default=argparse.SUPPRESS,
            help="Verbose logging",
        )
        p.add_argument(
            "-q", "--quiet", action="store_true", default=argparse.SUPPRESS, help="Quiet logging"
        )

    # GUI
    parser_gui = subparsers.add_parser("gui", help="Launch Windows 11 Desktop Application")
    add_common_flags(parser_gui)

    # Validate
    parser_validate = subparsers.add_parser(
        "validate", help="Validate a WebdynSunPM definition file"
    )
    add_common_flags(parser_validate)
    parser_validate.add_argument("input_file", help="Definition CSV to validate")
    parser_validate.add_argument(
        "--lenient",
        action="store_true",
        help="Lenient validation (allow address overlaps)",
    )

    # Extract Subparser
    parser_extract = subparsers.add_parser("extract", help="Extract registers from documentation")
    add_common_flags(parser_extract)
    parser_extract.add_argument("input_file", help="Source file (PDF/Excel/CSV/XML)")
    parser_extract.add_argument("-o", "--output", help="Output CSV")
    parser_extract.add_argument("--mapping", help="Mapping JSON")
    parser_extract.add_argument("--sheet", help="Excel sheet")
    parser_extract.add_argument("--pages", help="PDF pages")
    parser_extract.add_argument("--address-offset", type=int, default=0, help="Address offset")
    parser_extract.add_argument("--force", action="store_true", help="Force overwrite output file")

    # Generate
    parser_generate = subparsers.add_parser("generate", help="Generate definition from CSV")
    add_common_flags(parser_generate)
    parser_generate.add_argument("input_file", nargs="?", help="Input CSV")
    parser_generate.add_argument("-o", "--output", help="Output definition CSV")
    parser_generate.add_argument("--manufacturer", help="Manufacturer name")
    parser_generate.add_argument("--model", help="Model name")
    parser_generate.add_argument("--template", action="store_true", help="Generate template")
    parser_generate.add_argument(
        "--template-mode", choices=["input", "definition"], default="input"
    )
    parser_generate.add_argument("--protocol", default="modbusRTU")
    parser_generate.add_argument("--category", default="Inverter")
    parser_generate.add_argument("--forced-write", default="")
    parser_generate.add_argument("--address-offset", type=int, default=0, help="Address offset")
    parser_generate.add_argument("--report", help="Path to generate technical Markdown report")
    parser_generate.add_argument("--force", action="store_true", help="Force overwrite output file")

    # Run
    parser_run = subparsers.add_parser("run", help="Extract and Generate in one step")
    add_common_flags(parser_run)
    parser_run.add_argument("input_file", nargs="?", help="Source file (PDF/Excel/CSV/XML)")
    parser_run.add_argument("-o", "--output", help="Output definition CSV")
    parser_run.add_argument("--manufacturer", help="Manufacturer name")
    parser_run.add_argument("--model", help="Model name")
    parser_run.add_argument("--template", action="store_true", help="Generate template")
    parser_run.add_argument("--template-mode", choices=["input", "definition"], default="input")
    parser_run.add_argument("--mapping", help="Mapping JSON")
    parser_run.add_argument("--sheet", help="Excel sheet")
    parser_run.add_argument("--pages", help="PDF pages")
    parser_run.add_argument("--protocol", default="modbusRTU")
    parser_run.add_argument("--category", default="Inverter")
    parser_run.add_argument("--forced-write", default="")
    parser_run.add_argument("--address-offset", type=int, default=0, help="Address offset")
    parser_run.add_argument("--report", help="Path to generate technical Markdown report")
    parser_run.add_argument("--force", action="store_true", help="Force overwrite output file")
    parser_run.add_argument(
        "--no-validate", action="store_true", help="Skip post-generation validation"
    )

    # Report
    parser_report = subparsers.add_parser(
        "report", help="Generate comprehensive technical reports for definitions"
    )
    add_common_flags(parser_report)
    parser_report.add_argument(
        "target",
        nargs="?",
        default="tagged_definitions",
        help="Definition file or directory to generate report for (default: tagged_definitions)",
    )
    parser_report.add_argument("-o", "--output", help="Output Markdown report path")
    parser_report.add_argument("--json", help="Optional JSON metrics export path")

    # Tag
    parser_tag = subparsers.add_parser(
        "tag", help="Enrich definition CSV with official Webdyn SunPM tags"
    )
    add_common_flags(parser_tag)
    parser_tag.add_argument("input_file", help="Definition CSV file")
    parser_tag.add_argument("-o", "--output", help="Output definition CSV (defaults to in-place)")
    parser_tag.add_argument(
        "--category",
        default="",
        help="Equipment category (Inverter, Meter, Sensor, Battery, etc.)",
    )
    parser_tag.add_argument("--overwrite", action="store_true", help="Overwrite existing tags")
    parser_tag.add_argument(
        "--no-normalize",
        action="store_true",
        help="Do not normalize legacy tags to SunPM canonical standard",
    )

    # Audit
    parser_audit = subparsers.add_parser(
        "audit", help="Audit definition file(s) for errors, overlaps, tags, and units"
    )
    add_common_flags(parser_audit)
    parser_audit.add_argument(
        "target",
        nargs="?",
        default="Definitions",
        help="Definition file or directory to audit (default: Definitions)",
    )
    parser_audit.add_argument("--strict", action="store_true", help="Treat warnings as errors")
    parser_audit.add_argument(
        "--strict-overlap",
        action="store_true",
        help="Treat Modbus address overlaps as errors",
    )
    parser_audit.add_argument(
        "--fix",
        action="store_true",
        help="Automatically repair legacy tags, fill known tags, and remove duplicates",
    )
    parser_audit.add_argument(
        "-o", "--output", help="Output file when --fix is applied to a single file"
    )
    parser_audit.add_argument("--json", action="store_true", help="Output audit report as JSON")

    # Train
    parser_train = subparsers.add_parser(
        "train", help="Train semantic knowledge base from definition CSVs"
    )
    add_common_flags(parser_train)
    parser_train.add_argument(
        "--dir",
        default="Definitions",
        help="Directory containing definition CSV files (default: Definitions)",
    )
    parser_train.add_argument("--export-json", help="Export learned knowledge to a JSON file")
    parser_train.add_argument(
        "--summary",
        action="store_true",
        help="Display summary of learned metrics and tag distribution",
    )

    args = parser.parse_args(argv)

    if getattr(args, "gui", False) or args.command == "gui":
        from DefFileGenerator.gui import main as gui_main

        gui_main()
        sys.exit(0)

    if not args.command:
        parser.print_help()
        sys.exit(2)

    verbose = getattr(args, "verbose", False)
    quiet = getattr(args, "quiet", False)
    setup_logging(verbose=verbose, quiet=quiet)

    # Validate pages/sheet parameters depending on input file type
    input_file = getattr(args, "input_file", None)
    if input_file:
        ext = os.path.splitext(input_file)[1].lower()
        if getattr(args, "pages", None) and ext != ".pdf":
            logging.warning("--pages is only applicable for PDF files. Ignoring.")
        if getattr(args, "sheet", None) and ext not in (
            ".xlsx",
            ".xlsm",
            ".xltx",
            ".xltm",
        ):
            logging.warning("--sheet is only applicable for Excel files. Ignoring.")

    if args.command == "extract":
        extract_command(args)
    elif args.command == "validate":
        validate_command(args)
    elif args.command == "generate":
        generate_command(args)
    elif args.command == "run":
        run_command(args)
    elif args.command == "report":
        report_command(args)
    elif args.command == "tag":
        tag_command(args)
    elif args.command == "audit":
        audit_command(args)
    elif args.command == "train":
        train_command(args)


def main(args: list[str] | None = None) -> None:
    """
    Main entrypoint invoked when executing `deffilegen` or `python -m DefFileGenerator.main`.

    Args:
        args: Optional list of argument strings.
    """
    try:
        _run_cli(args)
    except KeyboardInterrupt:
        sys.exit(130)
    except SystemExit:
        raise
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
