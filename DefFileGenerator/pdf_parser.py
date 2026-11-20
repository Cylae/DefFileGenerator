"""
PDF Modbus Parser Module.

Extracts Modbus register mappings from PDF datasheets, manufacturer manuals, and technical specification guides.
Features automated table extraction with pdfplumber, regex text-based table fallbacks (for un-tabled/scanned layouts),
multiline header joining, page continuity tracking, and crash-resilient error handling.
"""

from __future__ import annotations

import io
import logging
import os
from collections.abc import Iterator
from pathlib import Path
from typing import Any

try:
    import pdfplumber

    HAS_PDFPLUMBER = True
except ImportError:
    pdfplumber = None  # type: ignore[assignment]
    HAS_PDFPLUMBER = False

from DefFileGenerator.extractor import (
    Extractor,
)


class PDFModbusParser:
    """
    Dedicated parser for extracting Modbus register maps from technical PDF documents.

    Features:
        - Structural table extraction across split pages
        - Fallback text pattern parser for plain-text tabular blocks
        - Multiline header assembly and alignment correction
        - Filtering of non-register tables (e.g. communication parameters, revision blocks)
        - Safe error recovery without crashing on malformed or partial PDFs
    """

    def __init__(self, mapping: dict[str, str] | None = None) -> None:
        """
        Initializes the PDFModbusParser with optional custom column mappings.

        Args:
            mapping: Optional dictionary mapping target Webdyn field names to source column names.
        """
        self.mapping = mapping
        self._extractor = Extractor(mapping=mapping)

    def get_page_count(self, filepath: str | Path) -> int:
        """
        Returns the total number of pages in the PDF document.

        Args:
            filepath: Path to the target PDF.

        Returns:
            int: Number of pages, or 0 if unreadable.
        """
        path_str = str(filepath)
        if not os.path.exists(path_str):
            logging.error(f"PDF file not found: {path_str}")
            return 0
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required to parse PDF files.")
            return 0

        try:
            with open(path_str, "rb") as f:
                with pdfplumber.open(io.BytesIO(f.read())) as pdf:
                    return len(pdf.pages)
        except Exception as exc:
            logging.warning(f"Failed to read page count from {path_str}: {exc}")
            return 0

    def extract_from_pdf(
        self,
        filepath: str | Path,
        pages: int | list[int | str] | str | None = None,
    ) -> Iterator[Iterator[dict[str, Any]]]:
        """
        Yields raw table row iterators from the PDF.

        Args:
            filepath: Path to the PDF document.
            pages: Optional page specification (e.g. 1, [1, 2], "3-5", "1,3,7").

        Returns:
            Iterator[Iterator[dict[str, Any]]]: Generator yielding per-table row dictionaries.
        """
        return self._extractor.extract_from_pdf(str(filepath), pages=pages)

    def extract_tables(
        self,
        filepath: str | Path,
        pages: int | list[int | str] | str | None = None,
    ) -> list[list[dict[str, Any]]]:
        """
        Extracts all identified Modbus tables into a list of materialized table lists.

        Args:
            filepath: Path to the PDF document.
            pages: Optional page numbers or range string.

        Returns:
            list[list[dict[str, Any]]]: List of tables, each represented as a list of row dicts.
        """
        tables: list[list[dict[str, Any]]] = []
        for table_gen in self.extract_from_pdf(filepath, pages=pages):
            try:
                rows = list(table_gen)
                if rows:
                    tables.append(rows)
            except Exception as exc:
                logging.warning(f"Error iterating table generator in {filepath}: {exc}")
        return tables

    def extract_and_clean(
        self,
        filepath: str | Path,
        pages: int | list[int | str] | str | None = None,
        address_offset: int = 0,
        manufacturer: str | None = None,
        category: str | None = None,
    ) -> list[dict[str, Any]]:
        """
        Extracts and normalizes Modbus registers from PDF document into standard Webdyn registers.

        Args:
            filepath: Path to the PDF document.
            pages: Optional page specification.
            address_offset: Integer offset to apply to addresses.
            manufacturer: Optional manufacturer override.
            category: Optional device category.

        Returns:
            list[dict[str, Any]]: Cleaned and normalized register dictionaries.
        """
        raw_tables = self.extract_from_pdf(filepath, pages=pages)
        cleaned = list(self._extractor.map_and_clean(raw_tables, address_offset=address_offset))
        if manufacturer or category:
            for r in cleaned:
                if manufacturer and "Manufacturer" not in r:
                    r["Manufacturer"] = manufacturer
                if category and "Category" not in r:
                    r["Category"] = category
        return cleaned


def parse_pdf_modbus(
    filepath: str | Path,
    pages: int | list[int | str] | str | None = None,
    address_offset: int = 0,
    manufacturer: str | None = None,
    category: str | None = None,
    mapping: dict[str, str] | None = None,
) -> list[dict[str, Any]]:
    """
    Convenience function to parse Modbus registers from a PDF file.

    Args:
        filepath: Path to the PDF file.
        pages: Page numbers or range string.
        address_offset: Address offset integer.
        manufacturer: Equipment manufacturer.
        category: Equipment category.
        mapping: Optional custom column mapping.

    Returns:
        list[dict[str, Any]]: Normalized register rows.
    """
    parser = PDFModbusParser(mapping=mapping)
    return parser.extract_and_clean(
        filepath=filepath,
        pages=pages,
        address_offset=address_offset,
        manufacturer=manufacturer,
        category=category,
    )
