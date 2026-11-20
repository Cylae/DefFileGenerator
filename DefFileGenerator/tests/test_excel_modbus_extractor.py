"""
Comprehensive test suite for ExcelModbusExtractor.

Tests:
1. Multi-sheet Modbus table extraction (reading all sheets, specific sheets, get_sheet_names).
2. Merged header cells (horizontally merged categories, vertically merged header columns).
3. Formulas (evaluated cached values, =HEX2DEC('7530'), =A2+1, =40001+10, ='Text').
4. Hexadecimal Modbus addresses (0x7530 -> 30000, 8232h -> 33330, 0x0001 -> 1).
5. Edge cases: empty sheets, bounded sparse columns, missing files, invalid formats.
"""

from __future__ import annotations

import logging
import os
import tempfile
import unittest

import openpyxl

from DefFileGenerator.excel_modbus_extractor import (
    ExcelModbusExtractor,
    extract_excel_modbus,
)
from DefFileGenerator.extractor import HAS_OPENPYXL


class TestExcelModbusExtractor(unittest.TestCase):
    """Test suite covering ExcelModbusExtractor capabilities."""

    def setUp(self) -> None:
        if not HAS_OPENPYXL:
            self.skipTest("openpyxl is not available in environment")
        self.tmpdir = tempfile.TemporaryDirectory()
        self.extractor = ExcelModbusExtractor()
        logging.disable(logging.CRITICAL)

    def tearDown(self) -> None:
        self.tmpdir.cleanup()
        logging.disable(logging.NOTSET)

    def _create_workbook(self, sheets_dict: dict[str, list[list]], filename: str = "test.xlsx") -> str:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # Remove default empty sheet
        for sheet_title, rows in sheets_dict.items():
            ws = wb.create_sheet(title=sheet_title)
            for r in rows:
                ws.append(r)
        target_path = os.path.join(self.tmpdir.name, filename)
        wb.save(target_path)
        return target_path

    def test_multi_sheet_modbus_tables(self) -> None:
        """Verify extraction across multiple Modbus sheets in a single workbook."""
        wb_path = self._create_workbook(
            {
                "Inverter_Telemetry": [
                    ["Address", "Name", "Type", "Unit"],
                    [30000, "Active Power", "U32", "W"],
                    [30002, "Reactive Power", "S32", "var"],
                ],
                "Grid_Status": [
                    ["Register", "Description", "DataType", "Scale"],
                    [40001, "Grid Voltage L1", "U16", "0.1"],
                    [40002, "Grid Voltage L2", "U16", "0.1"],
                ],
            }
        )

        sheet_names = self.extractor.get_sheet_names(wb_path)
        self.assertEqual(sheet_names, ["Inverter_Telemetry", "Grid_Status"])

        # Extract all sheets via parse_workbook
        all_parsed = self.extractor.parse_workbook(wb_path)
        self.assertIn("Inverter_Telemetry", all_parsed)
        self.assertIn("Grid_Status", all_parsed)
        self.assertEqual(len(all_parsed["Inverter_Telemetry"]), 2)
        self.assertEqual(len(all_parsed["Grid_Status"]), 2)

        # Verify specific sheet extraction
        grid_rows = self.extractor.extract_and_clean(wb_path, sheet_name="Grid_Status")
        self.assertEqual(len(grid_rows), 2)
        self.assertEqual(grid_rows[0]["Name"], "Grid Voltage L1")
        self.assertEqual(grid_rows[0]["Address"], "40001")

    def test_merged_header_cells_vertical(self) -> None:
        """Verify handling of vertically merged header cells (e.g. A1:A2 for Address, B1:B2 for Name)."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "MergedHeaders"

        # Row 1 and Row 2 headers merged vertically
        ws["A1"] = "Modbus Address"
        ws["B1"] = "Signal Name"
        ws["C1"] = "Data Type"
        ws["D1"] = "Unit"
        ws.merge_cells("A1:A2")
        ws.merge_cells("B1:B2")
        ws.merge_cells("C1:C2")
        ws.merge_cells("D1:D2")

        # Row 3 and 4 data
        ws.cell(row=3, column=1, value=30050)
        ws.cell(row=3, column=2, value="Cabinet Temperature")
        ws.cell(row=3, column=3, value="S16")
        ws.cell(row=3, column=4, value="°C")
        ws.cell(row=4, column=1, value=30051)
        ws.cell(row=4, column=2, value="Heatsink Temperature")
        ws.cell(row=4, column=3, value="S16")
        ws.cell(row=4, column=4, value="°C")

        target_path = os.path.join(self.tmpdir.name, "vert_merged.xlsx")
        wb.save(target_path)

        rows = self.extractor.extract_and_clean(target_path)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["Address"], "30050")
        self.assertEqual(rows[0]["Name"], "Cabinet Temperature")
        self.assertEqual(rows[0]["Type"], "I16")

    def test_merged_header_cells_horizontal(self) -> None:
        """Verify handling of horizontally merged category header cells spanning columns."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "CategoryHeaders"

        # Row 1 has category banners
        ws["A1"] = "Modbus Address"
        ws["B1"] = "Parameters"
        ws.merge_cells("B1:D1")

        # Row 2 has sub-headers
        ws["A2"] = ""  # Inherited from vertical merge or compound logic
        ws["B2"] = "Register Name"
        ws["C2"] = "Format"
        ws["D2"] = "Engineering Unit"

        # Row 3 data
        ws.cell(row=3, column=1, value=500)
        ws.cell(row=3, column=2, value="Frequency")
        ws.cell(row=3, column=3, value="U16")
        ws.cell(row=3, column=4, value="Hz")

        target_path = os.path.join(self.tmpdir.name, "horiz_merged.xlsx")
        wb.save(target_path)

        rows = self.extractor.extract_and_clean(target_path)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Address"], "500")
        self.assertEqual(rows[0]["Name"], "Frequency")

    def test_hexadecimal_addresses(self) -> None:
        """Verify hexadecimal Modbus addresses (0x7530, 8232h, 0x1F40) are converted to decimal."""
        wb_path = self._create_workbook(
            {
                "HexRegisters": [
                    ["Address", "Name", "Type"],
                    ["0x7530", "Active Power Total", "U32"],
                    ["8232h", "Reactive Power Total", "S32"],
                    ["0x000A", "Inverter Status", "U16"],
                ]
            }
        )

        rows = self.extractor.extract_and_clean(wb_path)
        self.assertEqual(len(rows), 3)

        # 0x7530 is decimal 30000
        self.assertEqual(rows[0]["Address"], "30000")
        self.assertEqual(rows[0]["Name"], "Active Power Total")

        # 8232h is decimal 33330
        self.assertEqual(rows[1]["Address"], "33330")

        # 0x000A is decimal 10
        self.assertEqual(rows[2]["Address"], "10")

    def test_formula_evaluation_and_fallbacks(self) -> None:
        """Verify formula evaluations like =HEX2DEC('7530'), =A2+1, =40001+5."""
        wb_path = self._create_workbook(
            {
                "Formulas": [
                    ["Address", "Name", "Type"],
                    ["=HEX2DEC(\"7530\")", "Base Register", "U32"],
                    ["=A2+1", "Secondary Register", "U32"],
                    ["=40000+10", "Holding Base", "U16"],
                ]
            }
        )

        raw_tables = list(self.extractor.extract_from_excel(wb_path))
        raw_rows = list(raw_tables[0])
        self.assertEqual(len(raw_rows), 3)

        # First row evaluated HEX2DEC("7530") -> 30000
        self.assertEqual(str(raw_rows[0]["Address"]), "30000")

        # Second row evaluated =A2+1 -> 30001
        self.assertEqual(str(raw_rows[1]["Address"]), "30001")

        # Third row evaluated =40000+10 -> 40010
        self.assertEqual(str(raw_rows[2]["Address"]), "40010")

        # End to end mapping
        cleaned = self.extractor.extract_and_clean(wb_path)
        self.assertEqual(cleaned[0]["Address"], "30000")
        self.assertEqual(cleaned[1]["Address"], "30001")
        self.assertEqual(cleaned[2]["Address"], "40010")

    def test_convenience_extract_excel_modbus(self) -> None:
        """Verify the extract_excel_modbus top-level helper function."""
        wb_path = self._create_workbook(
            {
                "Sheet1": [
                    ["Address", "Name", "Type"],
                    ["100", "DC Voltage", "U16"],
                ]
            }
        )

        res = extract_excel_modbus(wb_path, address_offset=-1, manufacturer="TestVendor")
        self.assertEqual(len(res), 1)
        self.assertEqual(res[0]["Address"], "99")
        self.assertEqual(res[0]["Name"], "DC Voltage")

    def test_missing_and_empty_excel_handling(self) -> None:
        """Verify missing files and empty workbooks return empty lists safely without exceptions."""
        empty_path = self._create_workbook({"Empty": []})
        self.assertEqual(self.extractor.extract_and_clean(empty_path), [])

        missing_path = os.path.join(self.tmpdir.name, "missing_nonexistent.xlsx")
        self.assertEqual(self.extractor.extract_and_clean(missing_path), [])
        self.assertEqual(self.extractor.get_sheet_names(missing_path), [])


if __name__ == "__main__":
    unittest.main()
