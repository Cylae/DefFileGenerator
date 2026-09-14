#!/usr/bin/env python3
"""
Wave 7 Coverage Expansion Tests.

Focuses on targeted unit test scenarios for previously uncovered branches
in extractor.py, def_gen.py, and web/app.py.
"""

import io
import os
import tempfile
import unittest
from unittest.mock import patch

from fastapi.testclient import TestClient

from DefFileGenerator.def_gen import (
    Generator,
    GeneratorConfig,
    run_generator,
)
from DefFileGenerator.extractor import Extractor
from web.app import app


class TestCoverageWave7Extractor(unittest.TestCase):
    def test_infer_table_columns(self):
        # Test table with valid register address and type columns
        table = [
            ["0x1000", "U16", "RO", "Test Reg 1", "1", "V", "1"],
            ["0x1001", "I32", "RW", "Test Reg 2", "2", "A", "10"],
        ]
        headers = Extractor._infer_table_columns(table)
        self.assertIsNotNone(headers)
        self.assertIn("Address", headers)
        self.assertIn("Type", headers)

        # Test empty or insufficient table
        self.assertIsNone(Extractor._infer_table_columns([]))
        self.assertIsNone(Extractor._infer_table_columns([["val1"]]))

    def test_extractor_missing_or_corrupted_excel(self):
        extractor = Extractor()
        # Test missing Excel file
        gen_list = list(extractor.extract_from_excel("non_existent_file.xlsx"))
        self.assertEqual(len(gen_list), 1)
        self.assertEqual(list(gen_list[0]), [])

        # Test corrupted Excel file
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(b"NOT A ZIP OR EXCEL FILE")
            tmp_path = tmp.name
        try:
            gen_list_corrupt = list(extractor.extract_from_excel(tmp_path))
            self.assertEqual(len(gen_list_corrupt), 1)
            self.assertEqual(list(gen_list_corrupt[0]), [])
        finally:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)

    def test_pdf_extraction_edge_cases(self):
        extractor = Extractor()
        # Non-existent PDF file
        gen_list = list(extractor.extract_from_pdf("non_existent.pdf"))
        self.assertEqual(gen_list, [])


class TestCoverageWave7DefGen(unittest.TestCase):
    def test_normalize_address_val_range(self):
        self.assertEqual(Generator.normalize_address_val("31657~31658"), "31657")
        self.assertEqual(Generator.normalize_address_val("0x8232 .. 0x82FE"), "33330")
        self.assertEqual(Generator.normalize_address_val("40001 - 40002"), "40001")

    def test_apply_address_offset_negative_warning(self):
        with self.assertLogs(level="WARNING") as cm:
            res = Generator.apply_address_offset("10", -20, line_num=5, name="TestVar")
            self.assertEqual(res, "-10")
            self.assertTrue(any("negative address" in log for log in cm.output))

    def test_calculate_coefficients(self):
        ca, cb = Generator._calculate_coefficients("1/10", "5.5", "2")
        self.assertEqual(ca, "10.000000")
        self.assertEqual(cb, "5.500000")

    def test_write_output_csv_header_configs(self):
        rows = [
            {
                "Info1": "3",
                "Info2": "40001",
                "Info3": "U16",
                "Info4": "",
                "Name": "Voltage",
                "Tag": "voltage",
                "CoefA": "1.000000",
                "CoefB": "0.000000",
                "Unit": "V",
                "Action": "4",
            }
        ]
        s_io = io.StringIO()
        # Test with positional string arguments fallback
        Generator.write_output_csv(s_io, rows, "ACME", "ModelX", "modbusTCP", "Inverter", "1")
        val = s_io.getvalue()
        self.assertIn("modbusTCP", val)
        self.assertIn("ACME", val)

    def test_run_generator_non_existent_input(self):
        cfg = GeneratorConfig(input_file="non_existent_input_file.csv", output="out.csv")
        with self.assertLogs(level="ERROR") as cm:
            run_generator(cfg)
            self.assertTrue(any("Input file not found" in log for log in cm.output))


class TestCoverageWave7Web(unittest.TestCase):
    def setUp(self):
        self.client = TestClient(app)

    def test_convert_file_exceeds_max_upload_size(self):
        # Patch UPLOAD_CHUNK_BYTES to 10 bytes and MAX_UPLOAD_BYTES to 20 bytes
        with patch("web.app.MAX_UPLOAD_BYTES", 20), patch("web.app.UPLOAD_CHUNK_BYTES", 5):
            files = {"file": ("large_file.csv", b"A" * 50, "text/csv")}
            data = {"manufacturer": "Mfg", "model": "Mod"}
            response = self.client.post("/api/convert", files=files, data=data)
            self.assertEqual(response.status_code, 413)
            self.assertIn("exceeds the", response.json()["detail"])

    def test_convert_file_no_filename_or_invalid_filename(self):
        files = {"file": ("foo.txt", b"header\n1", "text/plain")}
        response = self.client.post("/api/convert", files=files)
        self.assertEqual(response.status_code, 400)

    def test_validate_file_invalid_file(self):
        files = {"file": ("test.txt", b"invalid content", "text/plain")}
        response = self.client.post("/api/validate", files=files)
        self.assertEqual(response.status_code, 200)
        res_json = response.json()
        self.assertIn("valid", res_json)


if __name__ == "__main__":
    unittest.main()
