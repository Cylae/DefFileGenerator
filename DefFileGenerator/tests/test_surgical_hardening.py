import csv
import importlib.metadata
import logging
import os
import tempfile
import unittest
from unittest.mock import patch

from fastapi.testclient import TestClient

import web.app as web_app_module
from DefFileGenerator.def_gen import Generator


class TestPackagingContract(unittest.TestCase):
    def test_requires_python_matches_runtime_syntax_floor(self):
        dist = importlib.metadata.distribution("def-file-generator")
        self.assertEqual(dist.metadata["Requires-Python"], ">=3.10")


class TestBitfieldIntegrity(unittest.TestCase):
    def test_bits_must_fit_inside_one_modbus_register(self):
        self.assertTrue(Generator.validate_address("100_0_1", "BITS"))
        self.assertTrue(Generator.validate_address("100_15_1", "BITS"))
        self.assertFalse(Generator.validate_address("100_15_2", "BITS"))
        self.assertFalse(Generator.validate_address("100_16_1", "BITS"))
        self.assertFalse(Generator.validate_address("100_0_0", "BITS"))

    def test_strict_validation_detects_overlapping_bit_slices(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, "overlap.csv")
            with open(path, "w", newline="", encoding="utf-8") as handle:
                writer = csv.writer(handle, delimiter=";")
                writer.writerow(["modbusRTU", "Inverter", "M", "X", "", "", "", "", "", "", ""])
                writer.writerow(["1", "3", "100_0_4", "BITS", "", "Low nibble", "low", "1", "0", "", "4"])
                writer.writerow(["2", "3", "100_2_4", "BITS", "", "Overlap", "overlap", "1", "0", "", "4"])

            self.assertFalse(Generator().validate_csv(path, strict=True))

    def test_strict_validation_allows_disjoint_bit_slices(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, "disjoint.csv")
            with open(path, "w", newline="", encoding="utf-8") as handle:
                writer = csv.writer(handle, delimiter=";")
                writer.writerow(["modbusRTU", "Inverter", "M", "X", "", "", "", "", "", "", ""])
                writer.writerow(["1", "3", "100_0_4", "BITS", "", "Low nibble", "low", "1", "0", "", "4"])
                writer.writerow(["2", "3", "100_4_4", "BITS", "", "High nibble", "high", "1", "0", "", "4"])

            self.assertTrue(Generator().validate_csv(path, strict=True))


class TestAtomicOutput(unittest.TestCase):
    def test_failed_write_preserves_existing_output(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            output = os.path.join(tmpdir, "definition.csv")
            with open(output, "w", encoding="utf-8") as handle:
                handle.write("existing production content\n")

            rows = [
                {
                    "Info1": "3",
                    "Info2": "100",
                    "Info3": "U16",
                    "Info4": "",
                    "Name": "Good",
                    "Tag": "good",
                    "CoefA": "1.0",
                    "CoefB": "0.0",
                    "Unit": "V",
                    "Action": "4",
                },
                {"Info2": "101"},
            ]

            with self.assertRaises(KeyError):
                Generator.write_output_csv(output, rows, "M", "X")

            with open(output, encoding="utf-8") as handle:
                self.assertEqual(handle.read(), "existing production content\n")
            self.assertEqual(os.listdir(tmpdir), ["definition.csv"])


class TestWebHardening(unittest.TestCase):
    def setUp(self):
        self.client = TestClient(web_app_module.app)
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)

    def test_convert_rejects_oversized_upload(self):
        with patch.object(web_app_module, "MAX_UPLOAD_BYTES", 16):
            response = self.client.post(
                "/api/convert",
                files={"file": ("sample.csv", b"x" * 17, "text/csv")},
            )
        self.assertEqual(response.status_code, 413)
        self.assertIn("limit", response.json()["detail"].lower())

    def test_validate_rejects_oversized_upload(self):
        with patch.object(web_app_module, "MAX_UPLOAD_BYTES", 16):
            response = self.client.post(
                "/api/validate",
                files={"file": ("definition.csv", b"x" * 17, "text/csv")},
            )
        self.assertEqual(response.status_code, 413)

    def test_internal_extraction_exception_is_not_disclosed(self):
        with patch.object(
            web_app_module.Extractor,
            "extract_from_csv",
            side_effect=RuntimeError("secret /srv/private/customer/path"),
        ):
            response = self.client.post(
                "/api/convert",
                files={"file": ("sample.csv", b"Name,Address,Type\n", "text/csv")},
            )
        self.assertEqual(response.status_code, 400)
        detail = response.json()["detail"]
        self.assertEqual(detail, "Failed to extract registers from uploaded file.")
        self.assertNotIn("secret", detail)
        self.assertNotIn("/srv/private", detail)

    def test_wildcard_cors_does_not_advertise_credentials(self):
        response = self.client.options(
            "/api/health",
            headers={
                "Origin": "https://example.invalid",
                "Access-Control-Request-Method": "GET",
            },
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers.get("access-control-allow-origin"), "*")
        self.assertNotIn("access-control-allow-credentials", response.headers)


if __name__ == "__main__":
    unittest.main()
