"""
Excel Modbus Extractor Module.

Robust, high-performance extraction of Modbus register mappings from Excel workbooks (.xlsx, .xlsm, .xls).
Handles multi-sheet Modbus tables, merged header cells, formulas (evaluated or formula strings),
and hexadecimal register addresses (e.g. 0x7530, 8232h).
"""

from __future__ import annotations

import logging
import os
import zipfile
from collections.abc import Iterator
from pathlib import Path
from typing import Any

from DefFileGenerator.extractor import (
    HAS_OPENPYXL,
    ExtractionError,
    Extractor,
    openpyxl,
    xlrd,
)


class ExcelModbusExtractor:
    """
    Dedicated extractor for Modbus register definitions stored in Excel spreadsheets.

    Features:
        - Multi-sheet table discovery and extraction
        - Unmerging and propagation of merged header and metadata cells
        - Support for evaluated formulas and client-side formula evaluation
        - Hexadecimal address normalization (e.g. 0x7530 -> 30000, 8232h -> 33330)
        - Memory and cell-count bounded protection against ZipBomb and dimension inflation
    """

    def __init__(self, mapping: dict[str, str] | None = None) -> None:
        """
        Initializes the ExcelModbusExtractor with optional custom column mappings.

        Args:
            mapping: Optional dictionary mapping target Webdyn field names to source column names.
        """
        self.mapping = mapping
        self._extractor = Extractor(mapping=mapping)

    def get_sheet_names(self, filepath: str | Path) -> list[str]:
        """
        Returns the list of worksheet names contained within an Excel workbook.

        Args:
            filepath: Path to the Excel file (.xlsx, .xlsm, or .xls).

        Returns:
            list[str]: Sheet names in workbook order.
        """
        path_str = str(filepath)
        if not os.path.exists(path_str):
            logging.error(f"Excel file not found: {path_str}")
            return []

        if path_str.lower().endswith(".xls") and not zipfile.is_zipfile(path_str):
            if xlrd is None:
                raise ExtractionError("Legacy Excel format (.xls) requires 'xlrd'.")
            wb = xlrd.open_workbook(path_str)
            return wb.sheet_names()

        if not HAS_OPENPYXL:
            raise ExtractionError("openpyxl is required to read modern Excel workbooks.")

        wb = openpyxl.load_workbook(path_str, read_only=True)
        try:
            return list(wb.sheetnames)
        finally:
            wb.close()

    def extract_from_excel(
        self,
        filepath: str | Path,
        sheet_name: str | None = None,
    ) -> Iterator[Iterator[dict[str, Any]]]:
        """
        Yields an iterator of row dictionaries for each worksheet in the workbook.

        Args:
            filepath: Path to the target spreadsheet.
            sheet_name: Optional single sheet name to extract.

        Returns:
            Iterator[Iterator[dict[str, Any]]]: Generator yielding per-sheet row dictionaries.
        """
        return self._extractor.extract_from_excel(str(filepath), sheet_name=sheet_name)

    def extract_and_clean(
        self,
        filepath: str | Path,
        sheet_name: str | None = None,
        address_offset: int = 0,
        manufacturer: str | None = None,
        category: str | None = None,
    ) -> list[dict[str, Any]]:
        """
        Extracts, maps, and standardizes Modbus registers across worksheet(s) into Webdyn format.

        Args:
            filepath: Path to the Excel workbook.
            sheet_name: Optional specific worksheet to process.
            address_offset: Offset integer applied to base Modbus addresses.
            manufacturer: Optional manufacturer override.
            category: Optional device category (e.g. Inverter, Meter, Battery).

        Returns:
            list[dict[str, Any]]: List of cleaned, normalized register dictionaries.
        """
        raw_tables = self.extract_from_excel(filepath, sheet_name=sheet_name)
        cleaned = list(self._extractor.map_and_clean(raw_tables, address_offset=address_offset))
        if manufacturer or category:
            for r in cleaned:
                if manufacturer and "Manufacturer" not in r:
                    r["Manufacturer"] = manufacturer
                if category and "Category" not in r:
                    r["Category"] = category
        return cleaned

    def parse_workbook(
        self,
        filepath: str | Path,
        address_offset: int = 0,
        manufacturer: str | None = None,
        category: str | None = None,
    ) -> dict[str, list[dict[str, Any]]]:
        """
        Parses all sheets in an Excel workbook and returns registers keyed by sheet name.

        Args:
            filepath: Path to the Excel workbook.
            address_offset: Offset integer applied to base Modbus addresses.
            manufacturer: Optional manufacturer override.
            category: Optional device category.

        Returns:
            dict[str, list[dict[str, Any]]]: Sheet name mapped to normalized register records.
        """
        path_str = str(filepath)
        sheet_names = self.get_sheet_names(path_str)
        results: dict[str, list[dict[str, Any]]] = {}

        for sname in sheet_names:
            try:
                registers = self.extract_and_clean(
                    path_str,
                    sheet_name=sname,
                    address_offset=address_offset,
                    manufacturer=manufacturer,
                    category=category,
                )
                if registers:
                    results[sname] = registers
            except Exception as exc:
                logging.warning(f"Error parsing sheet '{sname}' in {path_str}: {exc}")

        return results


def extract_excel_modbus(
    filepath: str | Path,
    sheet_name: str | None = None,
    address_offset: int = 0,
    manufacturer: str | None = None,
    category: str | None = None,
    mapping: dict[str, str] | None = None,
) -> list[dict[str, Any]]:
    """
    Convenience function to extract and normalize Modbus registers from an Excel file.

    Args:
        filepath: Path to the Excel file.
        sheet_name: Optional worksheet name.
        address_offset: Address offset integer.
        manufacturer: Equipment manufacturer.
        category: Equipment category.
        mapping: Optional custom column mapping.

    Returns:
        list[dict[str, Any]]: Normalized register rows.
    """
    extractor = ExcelModbusExtractor(mapping=mapping)
    return extractor.extract_and_clean(
        filepath=filepath,
        sheet_name=sheet_name,
        address_offset=address_offset,
        manufacturer=manufacturer,
        category=category,
    )
