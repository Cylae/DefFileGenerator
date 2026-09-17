#!/usr/bin/env python3
"""
Format-Agnostic Modbus Register Extractor Module.

This module provides high-throughput, memory-bounded extraction of tabular Modbus
register maps from manufacturer documentation in diverse formats:
    - Microsoft Excel (.xlsx, .xlsm, .xltx, .xltm) via openpyxl
    - Adobe PDF (.pdf) via pdfplumber
    - Delimited text (.csv, .tsv) via Python csv module and sniffer
    - Extensible Markup Language (.xml) via defusedxml

Architecture & Pipeline:
    1. Stream Ingestion: Files are loaded into memory buffers or accessed as streams,
       preventing OS file descriptor lock contention on Windows systems.
    2. Smart Header Detection: Tables are scanned using keyword scoring to locate the
       true Modbus header row and merge multiline header banners. Non-register tables
       (e.g., communication parameters, document revision blocks) are filtered out.
    3. Inferred Fallback: When explicit headers are absent, column types are inferred
       statistically from cell value shapes (hex/decimal addresses, type codes, units).
    4. Two-Pass Heuristic Mapping: Target Webdyn fields (Address, Name, Type, etc.) are
       mapped to source column names using exact matching, synonym lookup, and fuzzy indicators.
    5. Normalization & Sanitization: Address whitespace is collapsed, register counts
       converted to byte lengths for strings, and gain divisors inverted to multipliers.
"""

from __future__ import annotations

import argparse
import csv
import io
import itertools
import json
import logging
import os
import re
import sys
import zipfile
from collections.abc import Iterable, Iterator
from typing import Any

# Named logger for this module
logger = logging.getLogger("DefFileGenerator.extractor")

# -----------------------------------------------------------------------------
# Optional Dependencies and Safe Parser Fallbacks
# -----------------------------------------------------------------------------

try:
    import openpyxl

    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

PDF_ERRORS: tuple[type[BaseException], ...]
try:
    import pdfplumber

    HAS_PDFPLUMBER = True
    try:
        from pdfminer.pdfparser import PDFSyntaxError
        from pdfplumber.utils.exceptions import PdfminerException

        PDF_ERRORS = (PDFSyntaxError, PdfminerException)
    except ImportError:
        PDF_ERRORS = (Exception,)
except ImportError:
    HAS_PDFPLUMBER = False
    PDF_ERRORS = (Exception,)

SECURITY_EXCEPTIONS: tuple[type[BaseException], ...]
try:
    from defusedxml import ElementTree as ET
    from defusedxml.common import (
        DefusedXmlException,
        DTDForbidden,
        EntitiesForbidden,
        ExternalReferenceForbidden,
    )

    HAS_DEFUSEDXML = True
    SECURITY_EXCEPTIONS = (
        DefusedXmlException,
        DTDForbidden,
        EntitiesForbidden,
        ExternalReferenceForbidden,
    )
except ImportError:
    HAS_DEFUSEDXML = False
    SECURITY_EXCEPTIONS = (Exception,)

XML_PARSE_ERRORS: tuple[type[BaseException], ...]
try:
    from defusedxml.ElementTree import ParseError as DefusedParseError

    XML_PARSE_ERRORS = (DefusedParseError,)
except ImportError:
    XML_PARSE_ERRORS = (Exception,)

# Import core generator utilities
try:
    from DefFileGenerator.def_gen import Generator, peek_generator
except ImportError:
    from .def_gen import Generator, peek_generator  # type: ignore[no-redef]

# -----------------------------------------------------------------------------
# Pre-compiled Regex Patterns for Extractor Performance
# -----------------------------------------------------------------------------

RE_CLEAN_WHITESPACE = re.compile(r"\s+")
RE_HEX_OR_DEC = re.compile(r"^(0x[0-9a-fA-F]+|\d{4,5})$")
RE_FIRST_ADDR_CANDIDATE = re.compile(r"^(0x[0-9a-fA-F]+|\d{1,5})$")
RE_HEURISTIC_TYPE = re.compile(r"^(u|i|s|f|str|bit|bitmap|bitfield)\s*\d*$", re.IGNORECASE)
RE_HEURISTIC_RW = re.compile(r"^(ro|rw|wo|r|w)$", re.IGNORECASE)
RE_HEURISTIC_UNIT = re.compile(r"^(v|a|w|kw|kwh|mwh|var|kvar|hz|%|deg|c|℃)$", re.IGNORECASE)
RE_COMPACT_DIGITS = re.compile(r"^\d+$")

# -----------------------------------------------------------------------------
# Domain Keyword Sets (Frozen for O(1) Fast Membership Checks)
# -----------------------------------------------------------------------------

MODBUS_HEADER_KEYWORDS = frozenset(
    {
        "address",
        "addr",
        "adresse",
        "register",
        "reg",
        "registre",
        "name",
        "description",
        "nom",
        "point",
        "type",
        "datatype",
        "format",
        "unit",
        "unité",
        "r/w",
        "rw",
        "action",
        "access",
        "accès",
        "scale",
        "factor",
        "gain",
        "offset",
        "length",
        "size",
        "res.",
        "start reg",
    }
)

PDF_METADATA_KEYWORDS = frozenset(
    {
        "prepared by",
        "firmware team",
        "doc no",
        "sheet:",
        "page of",
        "drawing no",
        "confidential",
        "all rights reserved",
    }
)

PDF_COMM_KEYWORDS = frozenset(
    {
        "baud rate",
        "parity",
        "data bits",
        "stop bits",
        "check (parity) bit",
        "interface",
        "subnet",
        "gateway",
        "tcp port",
        "port number",
    }
)

NAME_INDICATOR_TERMS = frozenset(
    {
        "description",
        "desc",
        "name",
        "nom",
        "designation",
        "parameter",
        "param",
        "variable",
        "signal",
        "label",
    }
)

NON_ADDRESS_COLUMN_TERMS = frozenset(
    {
        "country",
        "region",
        "standard",
        "category",
        "degree",
        "version",
        "date",
        "time",
        "team",
        "firmware",
    }
)


# -----------------------------------------------------------------------------
# Extractor Class
# -----------------------------------------------------------------------------


class Extractor:
    """
    Extracts, maps, and normalizes Modbus registers from various manufacturer formats.

    Attributes:
        mapping: Optional user-supplied dictionary mapping target fields to source column names.
        COLUMN_MAPPING: Comprehensive dictionary of standard synonyms for Modbus table headers.
    """

    COLUMN_MAPPING: dict[str, list[str]] = {
        "RegisterType": [
            "register type",
            "reg type",
            "modbus type",
            "registertype",
            "info1",
        ],
        "Address": [
            "address",
            "addr",
            "register",
            "reg",
            "index",
            "info2",
            "adresse",
            "start reg (dec)",
            "start reg",
            "start register",
            "start address",
            "registre",
            "adresse modbus",
            "registre modbus",
            "reg (dec)",
            "reg(dec)",
            "dec reg",
            "hex reg",
            "adresse (dec)",
            "modbus address",
            "modbus register",
            "register address",
        ],
        "Name": [
            "name",
            "description",
            "parameter",
            "variable",
            "signal",
            "signal name",
            "point",
            "point name",
            "designation",
            "nom",
            "grandeur",
            "quantity",
        ],
        "Type": [
            "data type",
            "datatype",
            "type",
            "format",
            "info3",
            "type de données",
            "type de donnees",
            "format de données",
            "data format",
        ],
        "Unit": ["unit", "units", "unité", "unite", "eng unit", "engineering unit"],
        "Tag": ["tag"],
        "Action": ["action", "access", "accès", "acces"],
        "ReadWrite": [
            "read/write",
            "read/ write",
            "read write",
            "r/w",
            "lecture/écriture",
            "lecture/ecriture",
            "l/e",
            "rd/wr",
        ],
        "Factor": [
            "scale",
            "factor",
            "multiplier",
            "ratio",
            "coefa",
            "coef_a",
            "res.",
            "resolution",
            "résolution",
            "accuracy",
            "precision",
            "facteur",
        ],
        "Gain": ["gain"],
        "Offset": ["offset", "bias", "coefficient b", "coefb", "coef_b"],
        "ScaleFactor": ["scalefactor", "scale factor"],
        "Length": [
            "length",
            "len",
            "size",
            "count",
            "quantity",
            "qty",
            "nb registers",
            "number of reg",
            "number of regs",
            "nb of reg",
            "no. of reg",
            "no of reg",
            "number of registers",
        ],
        "StartBit": ["startbit", "start bit", "bit offset", "start_bit"],
    }

    def __init__(self, mapping: dict[str, str] | None = None) -> None:
        """
        Initializes an Extractor instance.

        Args:
            mapping: Optional explicit target-to-source column name mapping dictionary.
        """
        self.mapping = mapping or {}

    @staticmethod
    def normalize_type(t: Any) -> str:
        """
        Normalizes a data type description using the Generator's normalization logic.

        Args:
            t: Raw data type description.

        Returns:
            str: Normalized data type code.
        """
        return Generator.normalize_type(t)

    @staticmethod
    def _infer_table_columns(table: list[list[Any]]) -> list[str] | None:
        """
        Infers header column roles by analyzing data patterns across table rows.

        Used when a PDF or Excel document has no header text row and starts
        immediately with register records.

        Args:
            table: 2D list of extracted table cells.

        Returns:
            Optional[list[str]]: Inferred header row strings, or None if detection fails.
        """
        if not table or len(table) < 1:
            return None
        cols_count = max(len(r) for r in table)
        col_scores = {
            c: {
                "addr": 0,
                "type": 0,
                "rw": 0,
                "len": 0,
                "name": 0,
                "unit": 0,
                "gain": 0,
            }
            for c in range(cols_count)
        }

        for row in table[:10]:
            for c, val in enumerate(row):
                if not val:
                    continue
                v_clean = RE_CLEAN_WHITESPACE.sub("", str(val).strip())
                if RE_HEX_OR_DEC.match(v_clean):
                    try:
                        val_int = int(v_clean, 0)
                        if 100 <= val_int <= 65535:
                            col_scores[c]["addr"] += 1
                    except ValueError:
                        pass
                if RE_HEURISTIC_TYPE.match(v_clean):
                    col_scores[c]["type"] += 1
                if RE_HEURISTIC_RW.match(v_clean):
                    col_scores[c]["rw"] += 1
                if v_clean in ("1", "2", "4", "8", "10", "15", "16", "32"):
                    col_scores[c]["len"] += 1
                if v_clean in ("1", "10", "100", "1000", "10000"):
                    col_scores[c]["gain"] += 1
                if RE_HEURISTIC_UNIT.match(v_clean):
                    col_scores[c]["unit"] += 1
                if len(str(val).strip()) > 3 and not RE_COMPACT_DIGITS.match(v_clean):
                    col_scores[c]["name"] += 1

        addr_col = max(col_scores.keys(), key=lambda c: col_scores[c]["addr"])
        type_col = max(col_scores.keys(), key=lambda c: col_scores[c]["type"])

        if (
            col_scores[addr_col]["addr"] >= 1
            and col_scores[type_col]["type"] >= 1
            and addr_col != type_col
        ):
            headers = [""] * cols_count
            headers[addr_col] = "Address"
            headers[type_col] = "Type"
            used = {addr_col, type_col}
            for attr, pat_name in [
                ("ReadWrite", "rw"),
                ("Name", "name"),
                ("Length", "len"),
                ("Unit", "unit"),
                ("Gain", "gain"),
            ]:
                cands = [c for c in col_scores if c not in used and col_scores[c][pat_name] >= 1]
                if cands:
                    best_c = max(cands, key=lambda c: col_scores[c][pat_name])
                    headers[best_c] = attr
                    used.add(best_c)
            return headers

        return None

    def extract_from_excel(
        self, filepath: str, sheet_name: str | None = None
    ) -> Iterator[Iterator[dict[str, Any]]]:
        """
        Extracts register tables from an Excel spreadsheet as lazy generator streams.

        Loads file into an in-memory BytesIO buffer to prevent Windows OS file lock issues.
        Performs smart header detection by scoring rows for Modbus keywords.

        Args:
            filepath: Path to the .xlsx / .xlsm spreadsheet.
            sheet_name: Optional single sheet name to extract. If None, processes all sheets.

        Yields:
            Iterator[dict[str, Any]]: Generator yielding dictionary rows for each sheet.
        """
        if not HAS_OPENPYXL:
            logging.error("openpyxl is required for Excel extraction.")
            return iter([])

        def excel_sheets_generator() -> Iterator[Iterator[dict[str, Any]]]:
            if not os.path.exists(filepath):

                def missing_file_gen():
                    logging.error(f"Excel file not found: {filepath}")
                    yield from ()

                yield missing_file_gen()
                return

            try:
                with open(filepath, "rb") as f:
                    file_data = f.read()
                wb = openpyxl.load_workbook(io.BytesIO(file_data), data_only=True)
            except (OSError, zipfile.BadZipFile, Exception) as exc:

                def io_err_gen(e=exc):
                    logging.error(f"File IO Error extracting from Excel {filepath}: {e}")
                    yield from ()

                yield io_err_gen()
                return

            sheet_names = [sheet_name] if sheet_name else wb.sheetnames

            for sname in sheet_names:

                def sheet_generator(name=sname) -> Iterator[dict[str, Any]]:
                    if name not in wb.sheetnames:
                        logging.error(f"Sheet '{name}' not found in {filepath}")
                        return
                    ws = wb[name]
                    rows = list(ws.iter_rows(values_only=True))
                    if not rows:
                        return

                    # Smart header detection: check up to first 15 rows for known Modbus columns
                    best_idx = 0
                    best_score = 0
                    for idx, r in enumerate(rows[:15]):
                        if not r:
                            continue
                        score = sum(
                            1
                            for c in r
                            if c is not None
                            and any(k in str(c).lower() for k in MODBUS_HEADER_KEYWORDS)
                        )
                        if score > best_score:
                            best_score = score
                            best_idx = idx

                    if best_score >= 2:
                        header_row = rows[best_idx]
                        data_rows = rows[best_idx + 1 :]
                    else:
                        first_non_empty = 0
                        for idx, r in enumerate(rows):
                            if any(cell is not None and str(cell).strip() for cell in r):
                                first_non_empty = idx
                                break
                        header_row = rows[first_non_empty]
                        data_rows = rows[first_non_empty + 1 :]

                    headers = [str(h).strip() if h is not None else "" for h in header_row]
                    for row in data_rows:
                        if any(cell is not None and str(cell).strip() for cell in row):
                            yield {
                                headers[i]: cell
                                for i, cell in enumerate(row)
                                if i < len(headers) and headers[i]
                            }

                yield sheet_generator()

        return excel_sheets_generator()

    def extract_from_pdf(
        self,
        filepath: str,
        pages: int | list[int | str] | str | None = None,
    ) -> Iterator[Iterator[dict[str, Any]]]:
        """
        Extracts register tables from a PDF document.

        Features:
            - Filters out metadata boxes and serial communication parameter tables.
            - Merges multiline headers across table splits.
            - Carries forward headers across consecutive page table continuations.
            - Corrects header alignment shifts between label and value columns.

        Args:
            filepath: Path to the PDF document.
            pages: Optional page numbers (1-indexed) as integer, list, or comma-separated string.

        Yields:
            Iterator[dict[str, Any]]: Generator yielding dictionary rows for each table.
        """
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return iter([])

        if not os.path.exists(filepath):
            logging.error(f"PDF file not found: {filepath}")
            return iter([])

        def pdf_tables_generator() -> Iterator[Iterator[dict[str, Any]]]:
            try:
                with open(filepath, "rb") as f:
                    file_data = f.read()
                pdf_stream = io.BytesIO(file_data)
                with pdfplumber.open(pdf_stream) as pdf:
                    target_pages = []
                    if pages is None:
                        target_pages = pdf.pages
                    else:
                        requested = pages if isinstance(pages, list) else [pages]
                        if isinstance(pages, str):
                            requested = [p.strip() for p in pages.split(",")]
                        for p in requested:
                            try:
                                idx = int(p) - 1
                                if 0 <= idx < len(pdf.pages):
                                    target_pages.append(pdf.pages[idx])
                                else:
                                    logging.warning(
                                        f"Page {p} is out of range (1-{len(pdf.pages)})"
                                    )
                            except (ValueError, TypeError):
                                logging.warning(f"Invalid page reference: {p}")

                    last_headers: list[str] | None = None

                    for page in target_pages:
                        tables = page.extract_tables()
                        for table in tables:
                            if not table or len(table) < 2:
                                continue

                            # Skip metadata header/footer boxes (document revision/author blocks)
                            if len(table) <= 3:
                                text_meta = " ".join(
                                    str(c) for row in table for c in row if c
                                ).lower()
                                if any(k in text_meta for k in PDF_METADATA_KEYWORDS):
                                    continue

                            # Skip serial communication / port configuration tables (2 columns)
                            if len(table[0]) <= 2:
                                text_comm = " ".join(
                                    str(c) for row in table for c in row if c
                                ).lower()
                                if any(k in text_comm for k in PDF_COMM_KEYWORDS):
                                    continue

                            def table_generator(
                                current_table=table,
                            ) -> Iterator[dict[str, Any]]:
                                nonlocal last_headers

                                # Scan first 8 rows to find the first candidate address
                                first_addr_idx = None
                                for idx, r in enumerate(current_table[:8]):
                                    if any(
                                        c
                                        and RE_FIRST_ADDR_CANDIDATE.match(
                                            RE_CLEAN_WHITESPACE.sub("", str(c).strip())
                                        )
                                        for c in r
                                    ):
                                        first_addr_idx = idx
                                        break

                                best_idx = 0
                                best_score = 0
                                max_scan = (
                                    first_addr_idx
                                    if first_addr_idx is not None
                                    else min(4, len(current_table))
                                )
                                if max_scan > 0:
                                    for idx, r in enumerate(current_table[:max_scan]):
                                        if not r:
                                            continue
                                        score = 0
                                        for c in r:
                                            if not c:
                                                continue
                                            c_str = str(c).lower().strip()
                                            c_clean = "".join(c_str.split())
                                            for k in MODBUS_HEADER_KEYWORDS:
                                                if k == "format" and "information" in c_str:
                                                    continue
                                                if k in c_str or k in c_clean:
                                                    score += 1
                                                    break
                                        if score > best_score:
                                            best_score = score
                                            best_idx = idx

                                cols_count = len(current_table[0])

                                if best_score >= 2:
                                    header_end = (
                                        first_addr_idx
                                        if (
                                            first_addr_idx is not None and first_addr_idx > best_idx
                                        )
                                        else best_idx + 1
                                    )
                                    combined_headers = []
                                    for c_i in range(cols_count):
                                        parts = []
                                        for r_i in range(best_idx, header_end):
                                            cell_val = current_table[r_i][c_i]
                                            if cell_val:
                                                text_cell = str(cell_val).replace("\n", " ").strip()
                                                if not (
                                                    "block" in text_cell.lower()
                                                    and "element" in text_cell.lower()
                                                ):
                                                    parts.append(text_cell)
                                        combined_headers.append(" ".join(parts).strip())
                                    data_rows = current_table[header_end:]
                                elif (
                                    last_headers
                                    and len(last_headers) == cols_count
                                    and first_addr_idx == 0
                                ):
                                    combined_headers = list(last_headers)
                                    data_rows = current_table
                                else:
                                    inferred = Extractor._infer_table_columns(current_table)
                                    if inferred:
                                        combined_headers = inferred
                                        data_rows = current_table
                                    else:
                                        combined_headers = [
                                            str(c).replace("\n", " ").strip() if c else ""
                                            for c in current_table[0]
                                        ]
                                        data_rows = current_table[1:]

                                # Header alignment shift correction
                                for c_i in range(len(combined_headers) - 1):
                                    if not combined_headers[c_i] and combined_headers[c_i + 1]:
                                        c_i_has_data = any(
                                            row[c_i] and str(row[c_i]).strip()
                                            for row in data_rows
                                            if len(row) > c_i
                                        )
                                        c_next_no_data = not any(
                                            row[c_i + 1] and str(row[c_i + 1]).strip()
                                            for row in data_rows
                                            if len(row) > c_i + 1
                                        )
                                        if c_i_has_data and c_next_no_data:
                                            combined_headers[c_i] = combined_headers[c_i + 1]
                                            combined_headers[c_i + 1] = ""

                                if best_score >= 2:
                                    last_headers = combined_headers

                                for row in data_rows:
                                    row_dict = {}
                                    for i, cell in enumerate(row):
                                        if i < len(combined_headers):
                                            row_dict[combined_headers[i]] = (
                                                str(cell).replace("\n", " ").strip() if cell else ""
                                            )
                                    if any(v.strip() for v in row_dict.values() if v):
                                        yield row_dict

                            yield table_generator()

            except (OSError, *PDF_ERRORS) as e:  # type: ignore[misc]
                logging.error(
                    f"File IO Error or PDF Syntax Error extracting from PDF {filepath}: {e}"
                )
            except (ValueError, TypeError, IndexError) as e:
                logging.error(f"Error extracting from PDF {filepath}: {e}")

        return pdf_tables_generator()

    def extract_from_csv(self, filepath: str) -> Iterator[Iterator[dict[str, Any]]]:
        """
        Extracts register tables from delimited CSV or TSV files.

        Auto-detects UTF-16, UTF-8-sig, or CP1252 character encodings and sniffs delimiters.
        Directly identifies already-formatted WebdynSunPM definition files.

        Args:
            filepath: Path to the CSV file.

        Yields:
            Iterator[dict[str, Any]]: Generator yielding dictionary rows.
        """

        def csv_tables_generator() -> Iterator[Iterator[dict[str, Any]]]:
            def csv_table_generator() -> Iterator[dict[str, Any]]:
                try:
                    with open(filepath, "rb") as f:
                        raw_bytes = f.read()

                    if raw_bytes.startswith((b"\xff\xfe", b"\xfe\xff")):
                        encoding = "utf-16"
                    else:
                        try:
                            raw_bytes.decode("utf-8")
                            encoding = "utf-8-sig"
                        except UnicodeDecodeError:
                            encoding = "cp1252"

                    text = raw_bytes.decode(encoding, errors="replace")
                    lines = text.splitlines()
                    if not lines:
                        return

                    first_line = lines[0].strip()
                    first_parts = [p.strip() for p in first_line.split(";")]
                    if len(first_parts) >= 4 and first_parts[0].lower().startswith("modbus"):
                        webdyn_headers = [
                            "Index",
                            "RegisterType",
                            "Address",
                            "Type",
                            "Info4",
                            "Name",
                            "Tag",
                            "Factor",
                            "Offset",
                            "Unit",
                            "Action",
                        ]
                        reader = csv.reader(lines[1:], delimiter=";")
                        for row in reader:
                            if len(row) >= 11 and any(cell.strip() for cell in row):
                                yield dict(zip(webdyn_headers, [c.strip() for c in row]))
                        return

                    snippet = text[:2048]
                    try:
                        dialect = csv.Sniffer().sniff(snippet, delimiters=";,\t")
                        delimiter = dialect.delimiter
                    except csv.Error:
                        delimiter = ","
                        for d in (";", ",", "\t"):
                            if d in snippet:
                                delimiter = d
                                break

                    f_io = io.StringIO(text)
                    dict_reader = csv.DictReader(f_io, delimiter=delimiter)
                    for d_row in dict_reader:
                        row_has_data = any(
                            (v.strip() if isinstance(v, str) else any(str(x).strip() for x in v))
                            for k, v in d_row.items()
                            if k is not None and v is not None
                        )
                        if row_has_data:
                            yield {str(k): v for k, v in d_row.items() if k is not None}
                except OSError as e:
                    logging.error(f"File IO Error extracting from CSV {filepath}: {e}")
                except csv.Error as e:
                    logging.error(f"CSV Parsing Error in {filepath}: {e}")
                except UnicodeError as e:
                    logging.error(f"Encoding Error extracting from CSV {filepath}: {e}")
                except (ValueError, TypeError, AttributeError) as e:
                    logging.error(f"Unexpected error extracting from CSV {filepath}: {e}")

            yield csv_table_generator()

        return csv_tables_generator()

    def extract_from_xml(self, filepath: str) -> Iterator[Iterator[dict[str, Any]]]:
        """
        Extracts register tables from XML documentation using defusedxml.

        Safely parses XML documents without entity expansion or external resource fetching.

        Args:
            filepath: Path to the XML file.

        Yields:
            Iterator[dict[str, Any]]: Generator yielding dictionary rows.
        """
        if not HAS_DEFUSEDXML:
            logging.error("defusedxml is required for secure XML parsing.")
            return iter([])

        def xml_tables_generator() -> Iterator[Iterator[dict[str, Any]]]:
            if not os.path.exists(filepath):

                def missing_xml_gen():
                    logging.error(f"XML file not found: {filepath}")
                    yield from ()

                yield missing_xml_gen()
                return

            def xml_generator() -> Iterator[dict[str, Any]]:
                try:
                    with open(filepath, "rb") as f:
                        tree = ET.parse(f)
                        root = tree.getroot()

                    seen = set()
                    for elem in root.iter():
                        row = {}
                        for child in elem:
                            if len(child) == 0 and child.text:
                                row[child.tag] = child.text.strip()
                        if len(row) >= 2:
                            js = json.dumps(row, sort_keys=True)
                            if js not in seen:
                                seen.add(js)
                                yield row
                except SECURITY_EXCEPTIONS as e:
                    logging.error(f"Security error parsing XML {filepath}: {e}")
                    raise
                except (OSError, *XML_PARSE_ERRORS) as e:  # type: ignore[misc]
                    logging.error(
                        f"File IO Error or Parsing Error extracting from XML {filepath}: {e}"
                    )
                except (ValueError, TypeError) as e:
                    logging.error(f"Error extracting from XML {filepath}: {e}")

            yield xml_generator()

        return xml_tables_generator()

    def map_and_clean(
        self,
        tables: Iterable[Iterable[dict[str, Any]]] | None,
        address_offset: int = 0,
    ) -> Iterator[dict[str, Any]]:
        """
        Resolves table column roles and transforms extracted raw dictionaries into
        standard intermediate register dictionaries.

        Pipeline per table:
            1. Buffers first 50 rows to discover all distinct column headers.
            2. Resolves mapping using 3-stage priority:
               - Exact target field name match (case-insensitive)
               - Exact synonym match from COLUMN_MAPPING
               - Substring indicator match (with non-address guards)
            3. Checks that an Address column was resolved.
            4. Normalizes rows:
               - Collapses fragmented address whitespace.
               - Multiplies string register length by 2 when expressed as words.
               - Inverts Gain divisors into multiplication Factors (1 / Gain).
               - Applies address offsets.

        Args:
            tables: Stream of extracted tables (each table is an iterable of row dicts).
            address_offset: Global numerical offset to shift register addresses.

        Yields:
            dict[str, Any]: Normalized register dictionary ready for definition generation.
        """
        if not tables:
            return

        for table in tables:
            has_rows, table = peek_generator(table)
            if not has_rows:
                continue
            iterator = iter(table)
            buffer = []
            try:
                for _ in range(50):
                    buffer.append(next(iterator))
            except StopIteration:
                pass
            if not buffer:
                continue

            all_keys = list(
                dict.fromkeys(
                    str(col)
                    for r in buffer
                    for col in r.keys()
                    if col is not None and str(col).strip() != ""
                )
            )

            col_map = {}
            used_src_cols = set()

            # Apply explicit user mapping first
            for target, source in self.mapping.items():
                if source in all_keys:
                    col_map[target] = source
                    used_src_cols.add(source)

            detection_order = [
                "RegisterType",
                "Address",
                "Name",
                "Type",
                "Unit",
                "Action",
                "ReadWrite",
                "Tag",
                "Factor",
                "Gain",
                "Offset",
                "ScaleFactor",
                "Length",
                "StartBit",
            ]

            # Stage 1: Exact target-name match (case-insensitive and space-normalized)
            for target in detection_order:
                if target in col_map:
                    continue
                target_low = target.lower()
                for src_col in all_keys:
                    if src_col in used_src_cols:
                        continue
                    src_str = str(src_col).lower().strip()
                    src_clean = RE_CLEAN_WHITESPACE.sub("", src_str)
                    if src_str == target_low or src_clean == target_low:
                        col_map[target] = src_col
                        used_src_cols.add(src_col)
                        break

            # Stage 2: Pattern-based exact match from COLUMN_MAPPING
            for target in detection_order:
                if target in col_map:
                    continue
                patterns = self.COLUMN_MAPPING.get(target, [target.lower()])
                clean_patterns = [RE_CLEAN_WHITESPACE.sub("", p) for p in patterns]
                for src_col in all_keys:
                    if src_col in used_src_cols:
                        continue
                    src_str = str(src_col).lower().strip()
                    src_clean = RE_CLEAN_WHITESPACE.sub("", src_str)
                    if src_str in patterns or src_clean in clean_patterns:
                        col_map[target] = src_col
                        used_src_cols.add(src_col)
                        break

            # Stage 3: Partial substring matches with non-address guards
            for target in detection_order:
                if target in col_map:
                    continue
                patterns = self.COLUMN_MAPPING.get(target, [target.lower()])
                clean_patterns = [RE_CLEAN_WHITESPACE.sub("", p) for p in patterns]
                for src_col in all_keys:
                    if src_col in used_src_cols:
                        continue
                    src_low = str(src_col).lower().strip()
                    src_clean = RE_CLEAN_WHITESPACE.sub("", src_low)

                    # Guard: If column contains name/description terms or metadata, do NOT map as Address
                    if target == "Address" and (
                        any(ind in src_low for ind in NAME_INDICATOR_TERMS)
                        or any(ind in src_clean for ind in NAME_INDICATOR_TERMS)
                        or any(t in src_low for t in NON_ADDRESS_COLUMN_TERMS)
                        or any(t in src_clean for t in NON_ADDRESS_COLUMN_TERMS)
                    ):
                        continue

                    if any(p in src_low for p in patterns) or any(
                        cp in src_clean for cp in clean_patterns
                    ):
                        col_map[target] = src_col
                        used_src_cols.add(src_col)
                        break

            if "Address" not in col_map:
                continue

            length_src = str(col_map.get("Length", "")).lower()
            length_is_register_quantity = any(
                p in length_src for p in ("quantity", "count", "qty", "reg")
            )

            def process_row(
                r: dict[str, Any],
                col_map=col_map,
                length_is_register_quantity=length_is_register_quantity,
            ) -> dict[str, Any] | None:
                new_row = {target: r.get(src_col) for target, src_col in col_map.items()}
                addr_val = str(new_row.get("Address") or "").strip()
                if not addr_val:
                    return None

                # Collapse whitespace in addresses caused by fragmented font rendering in PDFs
                compact_addr = RE_CLEAN_WHITESPACE.sub("", addr_val)
                if compact_addr.isdigit() or (
                    compact_addr.lower().startswith("0x") and compact_addr[2:].isalnum()
                ):
                    addr_val = compact_addr

                sbit = str(r.get(col_map.get("StartBit", ""), "")).strip()
                slen = str(r.get(col_map.get("Length", ""), "")).strip()
                compact_slen = RE_CLEAN_WHITESPACE.sub("", slen)
                if compact_slen.isdigit():
                    slen = compact_slen

                raw_type = new_row.get("Type", "U16")
                dtype = Generator.normalize_type(raw_type)
                new_row["Type"] = dtype
                addr = addr_val

                is_string_type = dtype == "STRING" or dtype.startswith("STR")
                if is_string_type and slen.isdigit() and length_is_register_quantity:
                    # Convert 16-bit word length into Webdyn byte length (1 register = 2 bytes)
                    slen = str(int(slen) * 2)

                if dtype == "BITS" and "_" not in addr:
                    if sbit != "":
                        if slen == "":
                            slen = "1"
                    else:
                        sbit = "0"
                        if slen == "":
                            slen = "16"
                    addr = f"{addr}_{sbit}_{slen}"
                elif is_string_type and slen != "" and "_" not in addr:
                    addr = f"{addr}_{slen}"

                new_row["Address"] = Generator.apply_address_offset(addr, address_offset)

                # Gain column expresses a divisor (e.g. Huawei 10 -> 0.1 multiplier)
                if not new_row.get("Factor"):
                    gain_str = str(new_row.get("Gain", "")).strip()
                    compact_gain = RE_CLEAN_WHITESPACE.sub("", gain_str)
                    if compact_gain.isdigit():
                        gain_str = compact_gain
                    if gain_str:
                        gain_val = Generator._parse_numeric(gain_str, default=0.0)
                        if gain_val:
                            new_row["Factor"] = f"{1.0 / gain_val:.6f}"

                if new_row.get("Factor") is not None:
                    new_row["Factor"] = str(Generator._parse_numeric(new_row["Factor"], 1.0))

                # Fallback to ReadWrite column if Action was not directly provided
                if not new_row.get("Action"):
                    rw_str = str(new_row.get("ReadWrite", "")).strip()
                    if rw_str:
                        new_row["Action"] = rw_str

                if not new_row.get("RegisterType"):
                    new_row["RegisterType"] = "Holding Register"

                return new_row

            for row in itertools.chain(buffer, iterator):
                processed = process_row(row)
                if processed:
                    yield processed


def main() -> None:
    """Command-line entrypoint for the standalone extractor script."""
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    parser = argparse.ArgumentParser(description="Extract register information from documentation.")
    parser.add_argument("input_file", help="Path to input document (PDF, Excel, CSV, XML)")
    parser.add_argument("-o", "--output", help="Output simplified CSV destination")
    parser.add_argument("--mapping", help="Path to JSON column mapping file")
    parser.add_argument("--sheet", help="Target Excel sheet name")
    parser.add_argument("--pages", help="Target PDF page numbers (e.g. 1,2,5-8)")
    parser.add_argument(
        "--address-offset",
        type=int,
        default=0,
        help="Global address offset shift",
    )
    args = parser.parse_args()

    mapping = {}
    if args.mapping:
        with open(args.mapping) as f:
            mapping = json.load(f)
    extractor = Extractor(mapping)
    ext = os.path.splitext(args.input_file)[1].lower()
    pages = args.pages

    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        raw = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == ".pdf":
        raw = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == ".csv":
        raw = extractor.extract_from_csv(args.input_file)
    elif ext == ".xml":
        raw = extractor.extract_from_xml(args.input_file)
    else:
        logging.error(f"Unsupported extension: {ext}")
        sys.exit(1)

    mapped = list(extractor.map_and_clean(raw, args.address_offset))

    out = open(args.output, "w", newline="", encoding="utf-8") if args.output else sys.stdout
    writer = csv.DictWriter(
        out,
        fieldnames=[
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
        ],
        extrasaction="ignore",
    )
    writer.writeheader()
    writer.writerows(mapped)
    if args.output:
        out.close()


if __name__ == "__main__":
    main()
