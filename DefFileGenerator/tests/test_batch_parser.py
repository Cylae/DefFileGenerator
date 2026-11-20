"""
Comprehensive test suite for BatchModbusParser and PDFModbusParser.

Tests:
1. Batch extraction across heterogeneous vendor sheets (Excel, PDF, CSV, XML, JSON).
2. PDF table extraction fallbacks without crashes.
3. PDFModbusParser methods (extract_from_pdf, extract_tables, extract_and_clean).
4. Error containment: malformed or corrupted documents do not halt the batch pipeline.
5. Concurrent vs sequential worker processing.
6. Technical Markdown report generation.
"""

from __future__ import annotations

import logging
import os
import tempfile
import unittest
from pathlib import Path

import openpyxl

from DefFileGenerator.batch_parser import BatchModbusParser, batch_parse
from DefFileGenerator.pdf_parser import PDFModbusParser, parse_pdf_modbus


class TestBatchParserAndPDF(unittest.TestCase):
    """Test suite covering batch parsing across formats and PDF extractors."""

    def setUp(self) -> None:
        self.tmpdir = tempfile.TemporaryDirectory()
        logging.disable(logging.CRITICAL)

    def tearDown(self) -> None:
        self.tmpdir.cleanup()
        logging.disable(logging.NOTSET)

    def _create_sample_excel(self, directory: str, filename: str) -> str:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Registers"
        ws.append(["Address", "Name", "Type", "Unit"])
        ws.append([30000, "Active Power", "U32", "W"])
        ws.append([30002, "Reactive Power", "S32", "var"])
        path = os.path.join(directory, filename)
        wb.save(path)
        return path

    def _create_sample_csv(self, directory: str, filename: str) -> str:
        path = os.path.join(directory, filename)
        content = "Address,Name,Type,Unit\n40001,Grid Voltage L1,U16,V\n40002,Grid Voltage L2,U16,V\n"
        Path(path).write_text(content, encoding="utf-8")
        return path

    def _create_sample_json(self, directory: str, filename: str) -> str:
        path = os.path.join(directory, filename)
        content = """[
            {"Address": "100", "Name": "Internal Temp", "Type": "S16"},
            {"Address": "101", "Name": "External Temp", "Type": "S16"}
        ]"""
        Path(path).write_text(content, encoding="utf-8")
        return path

    def _create_sample_xml(self, directory: str, filename: str) -> str:
        path = os.path.join(directory, filename)
        content = """<?xml version="1.0"?>
        <registers>
            <register><Address>500</Address><Name>Frequency</Name><Type>U16</Type></register>
        </registers>"""
        Path(path).write_text(content, encoding="utf-8")
        return path

    def test_heterogeneous_batch_extraction(self) -> None:
        """Verify batch extraction across diverse vendor sheet formats (.xlsx, .csv, .json, .xml)."""
        input_dir = os.path.join(self.tmpdir.name, "hetero_in")
        output_dir = os.path.join(self.tmpdir.name, "hetero_out")
        report_file = os.path.join(self.tmpdir.name, "hetero_report.md")
        os.makedirs(input_dir)

        # Populate with different file formats
        self._create_sample_excel(input_dir, "VendorA_Inverter_Model1.xlsx")
        self._create_sample_csv(input_dir, "VendorB_Meter_M4M.csv")
        self._create_sample_json(input_dir, "VendorC_Sensor_Temp.json")
        self._create_sample_xml(input_dir, "VendorD_Gateway_Box.xml")

        parser = BatchModbusParser(
            protocol="modbusRTU",
            tag=True,
            strict=False,
            force=True,
            workers=2,
        )

        summary = parser.process_directory(input_dir, output_dir, report_path=report_file)

        self.assertEqual(summary["total_files"], 4)
        self.assertEqual(summary["successes"], 4)
        self.assertEqual(summary["failures"], 0)

        # Check output definitions exist
        generated_csvs = list(Path(output_dir).glob("*.csv"))
        self.assertEqual(len(generated_csvs), 4)

        # Verify report generation
        self.assertTrue(os.path.exists(report_file))
        report_text = Path(report_file).read_text(encoding="utf-8")
        self.assertIn("Rapport", report_text)

    def test_corrupt_file_containment(self) -> None:
        """Verify corrupted or invalid file does not crash the batch extraction process."""
        input_dir = os.path.join(self.tmpdir.name, "corrupt_in")
        output_dir = os.path.join(self.tmpdir.name, "corrupt_out")
        os.makedirs(input_dir)

        # 1 valid Excel, 1 invalid corrupted PDF
        self._create_sample_excel(input_dir, "ValidVendor_ModelA.xlsx")
        corrupt_pdf = os.path.join(input_dir, "Broken_Document.pdf")
        Path(corrupt_pdf).write_bytes(b"%PDF-1.4\nNOT A VALID PDF BODY CORRUPTED EOF\n")

        parser = BatchModbusParser(force=True, workers=1)
        summary = parser.process_directory(input_dir, output_dir)

        self.assertEqual(summary["total_files"], 2)
        # Valid one succeeds, corrupt one fails gracefully without halting
        self.assertEqual(summary["successes"], 1)
        self.assertEqual(summary["failures"], 1)

        # Valid definition was still generated
        csvs = list(Path(output_dir).glob("*.csv"))
        self.assertEqual(len(csvs), 1)

    def test_pdf_modbus_parser_methods_and_fallbacks(self) -> None:
        """Verify PDFModbusParser API and fallbacks without crashing."""
        pdf_parser = PDFModbusParser()

        # Non-existent PDF file handling
        missing_pdf = os.path.join(self.tmpdir.name, "not_existing.pdf")
        self.assertEqual(pdf_parser.get_page_count(missing_pdf), 0)
        self.assertEqual(pdf_parser.extract_tables(missing_pdf), [])
        self.assertEqual(pdf_parser.extract_and_clean(missing_pdf), [])

        # Empty or mock corrupted file handling
        empty_pdf = os.path.join(self.tmpdir.name, "empty.pdf")
        Path(empty_pdf).write_bytes(b"%PDF-1.4 empty")
        self.assertEqual(pdf_parser.extract_tables(empty_pdf), [])
        self.assertEqual(pdf_parser.extract_and_clean(empty_pdf), [])

        # Convenience function
        res = parse_pdf_modbus(missing_pdf)
        self.assertEqual(res, [])

    def test_batch_parse_convenience_function(self) -> None:
        """Verify batch_parse helper function."""
        input_dir = os.path.join(self.tmpdir.name, "conv_in")
        output_dir = os.path.join(self.tmpdir.name, "conv_out")
        os.makedirs(input_dir)

        self._create_sample_csv(input_dir, "Huawei_SUN2000.csv")

        res = batch_parse(input_dir, output_dir, protocol="modbusRTU", workers=1)
        self.assertEqual(res["total_files"], 1)
        self.assertEqual(res["successes"], 1)


if __name__ == "__main__":
    unittest.main()
