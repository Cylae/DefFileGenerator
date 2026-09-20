#!/usr/bin/env python3
"""
Independent test suite expansion for autonomous hardening and edge-case verification.
"""

import unittest

from fastapi.testclient import TestClient

from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor
from web.app import app


class TestAutonomousHardening(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.extractor = Extractor()

    def test_normalize_address_val_hex_and_negative(self):
        self.assertEqual(Generator.normalize_address_val("-0x10"), "-16")
        self.assertEqual(Generator.normalize_address_val("-10h"), "-16")
        self.assertEqual(Generator.normalize_address_val("-0X1A"), "-26")
        self.assertEqual(Generator.normalize_address_val("0x1A"), "26")
        self.assertEqual(Generator.normalize_address_val("100h"), "256")

    def test_sanitize_csv_field_surrogate_and_formula(self):
        # Surrogate character handling
        raw_surrogate = "test\ud800field"
        sanit = Generator.sanitize_csv_field(raw_surrogate)
        self.assertNotIn("\ud800", sanit)

        # Formula injection characters
        self.assertTrue(Generator.sanitize_csv_field("=1+1").startswith("'"))
        self.assertTrue(Generator.sanitize_csv_field("+CMD").startswith("'"))
        self.assertTrue(Generator.sanitize_csv_field("-CALC").startswith("'"))

    def test_extractor_map_and_clean_non_string_keys(self):
        table = [
            [{None: "invalid", "": "blank", "Address": "40001", "Name": "TestReg", "Type": "U16"}]
        ]
        mapped = list(self.extractor.map_and_clean(table))
        self.assertEqual(len(mapped), 1)
        self.assertEqual(mapped[0]["Address"], "40001")
        self.assertEqual(mapped[0]["Name"], "TestReg")

    def test_web_api_convert_and_validate_edge_cases(self):
        client = TestClient(app)

        # Test empty file upload in /api/convert
        res = client.post(
            "/api/convert",
            files={"file": ("  ", b"", "text/csv")},
            data={"manufacturer": "Test", "model": "Test"},
        )
        self.assertEqual(res.status_code, 400)

        # Test empty file upload in /api/validate
        res_v = client.post(
            "/api/validate",
            files={"file": ("   ", b"", "text/csv")},
        )
        self.assertEqual(res_v.status_code, 400)

    def test_web_convert_unhandled_exception_handling(self):
        client = TestClient(app)
        # Upload a CSV that causes an unhandled error inside run_generator by monkeypatching or sending unexpected input
        from unittest.mock import patch

        csv_data = b"Address,Name,Type\n40001,Test,U16\n"
        with patch("web.app.run_generator", side_effect=RuntimeError("Simulated engine failure")):
            res = client.post(
                "/api/convert",
                files={"file": ("test.csv", csv_data, "text/csv")},
                data={"manufacturer": "Test", "model": "Test"},
            )
            self.assertEqual(res.status_code, 400)
            self.assertIn("Core generator processing failed", res.json()["detail"])

    def test_validate_address_extreme_hex_and_octal(self):
        self.assertEqual(Generator.normalize_address_val("0x0"), "0")
        self.assertEqual(Generator.normalize_address_val("0x10000"), "65536")
        self.assertEqual(Generator.normalize_address_val("0o777"), "511")

    def test_sanitize_csv_field_unicode_control_characters(self):
        raw = "\ud800=1+1"
        sanit = Generator.sanitize_csv_field(raw)
        self.assertNotIn("\ud800", sanit)
        self.assertTrue(sanit.startswith("'"))


if __name__ == "__main__":
    unittest.main()
