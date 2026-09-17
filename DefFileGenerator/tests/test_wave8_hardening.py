import io
import os
import tempfile
import unittest

from fastapi.testclient import TestClient

from DefFileGenerator import CSVHeaderConfig, RegisterEntry
from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor
from web.app import app


class TestWave8Hardening(unittest.TestCase):
    def setUp(self):
        self.client = TestClient(app)
        self.tmpdir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.tmpdir.cleanup()

    def test_register_entry_dataclass_contract(self):
        entry = RegisterEntry(
            info1="3", address="40001", dtype="U16", name="ActivePower", line_num=10
        )
        self.assertEqual(entry.info1, "3")
        self.assertEqual(entry.address, "40001")
        self.assertEqual(entry.dtype, "U16")
        self.assertEqual(entry.name, "ActivePower")
        self.assertEqual(entry.line_num, 10)

        # Verify exported at package level
        import DefFileGenerator

        self.assertIn("RegisterEntry", DefFileGenerator.__all__)
        self.assertIn("CSVHeaderConfig", DefFileGenerator.__all__)

    def test_check_address_overlap_dual_calling_convention(self):
        gen = Generator()
        address_usage = {}
        warned_lines = set()

        # 1. Calling with RegisterEntry dataclass
        entry1 = RegisterEntry("3", "40001", "U32", "Var1", 1)
        overlap1 = gen._check_address_overlap(entry1, address_usage, warned_lines)
        self.assertFalse(overlap1)

        # Overlapping entry
        entry2 = RegisterEntry("3", "40002", "U16", "Var2", 2)
        overlap2 = gen._check_address_overlap(entry2, address_usage, warned_lines)
        self.assertTrue(overlap2)
        self.assertIn((1, 2), warned_lines)

        # 2. Calling with legacy positional arguments
        address_usage_legacy = {}
        warned_lines_legacy = set()
        overlap_leg1 = gen._check_address_overlap(
            "3", "40001", "U32", "Var1", 1, address_usage_legacy, warned_lines_legacy
        )
        self.assertFalse(overlap_leg1)

        overlap_leg2 = gen._check_address_overlap(
            "3", "40002", "U16", "Var2", 2, address_usage_legacy, warned_lines_legacy
        )
        self.assertTrue(overlap_leg2)
        self.assertIn((1, 2), warned_lines_legacy)

    def test_validate_address_string_zero_or_negative_length(self):
        gen = Generator()
        # Valid string address
        self.assertTrue(gen.validate_address("30001_10", "STRING"))
        # Invalid: 0 length
        self.assertFalse(gen.validate_address("30001_0", "STRING"))
        # Invalid format
        self.assertFalse(gen.validate_address("30001_-5", "STRING"))

    def test_parse_numeric_multi_slash(self):
        gen = Generator()
        self.assertEqual(gen._parse_numeric("1/10", default=0.0), 0.1)
        self.assertEqual(gen._parse_numeric("1/2/3", default=0.0), 0.0)
        self.assertEqual(gen._parse_numeric("1/0", default=0.0), 0.0)
        self.assertEqual(gen._parse_numeric("/5", default=0.0), 0.0)
        self.assertEqual(gen._parse_numeric("5/", default=0.0), 0.0)

    def test_atomic_write_missing_directory_handled_cleanly(self):
        gen = Generator()
        nested_out = os.path.join(self.tmpdir.name, "deep", "nested", "path", "out.csv")
        self.assertFalse(os.path.exists(os.path.dirname(nested_out)))

        rows = [
            {
                "Info1": "3",
                "Info2": "40001",
                "Info3": "U16",
                "Info4": "",
                "Name": "Test",
                "Tag": "test",
                "CoefA": "1.000000",
                "CoefB": "0.000000",
                "Unit": "V",
                "Action": "4",
            }
        ]
        hdr = CSVHeaderConfig(manufacturer="Mfg", model="Mod")
        # Should gracefully catch OSError and log error without uncaught exception
        gen.write_output_csv(nested_out, rows, hdr)
        self.assertFalse(os.path.exists(nested_out))

    def test_extract_pdf_missing_file_clean_handling(self):
        extractor = Extractor()
        missing_path = os.path.join(self.tmpdir.name, "does_not_exist.pdf")
        tables = list(extractor.extract_from_pdf(missing_path))
        self.assertEqual(len(tables), 0)

    def test_extract_pdf_empty_file_clean_handling(self):
        extractor = Extractor()
        empty_path = os.path.join(self.tmpdir.name, "empty.pdf")
        with open(empty_path, "wb") as f:
            f.write(b"")
        tables = list(extractor.extract_from_pdf(empty_path))
        self.assertEqual(len(tables), 0)

    def test_web_convert_hostile_filename_characters(self):
        csv_content = "Register,Name,Data Type,Unit,Scale,Access\n40001,Power,uint16,W,1,R\n"
        file_obj = io.BytesIO(csv_content.encode("utf-8"))
        response = self.client.post(
            "/api/convert",
            files={"file": ("..\\..\\nested/hack.csv", file_obj, "text/csv")},
            data={"manufacturer": "VeryLong" * 20, "model": "SuperModel" * 20},
        )
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertTrue(data["success"])
        # Filename should be safely bounded
        self.assertTrue(len(data["filename"]) < 150)
        self.assertTrue(data["filename"].endswith("_definition.csv"))

    def test_web_convert_dots_only_filename_rejected(self):
        file_obj = io.BytesIO(b"data")
        response = self.client.post(
            "/api/convert",
            files={"file": ("..", file_obj, "text/csv")},
        )
        self.assertEqual(response.status_code, 400)


if __name__ == "__main__":
    unittest.main()
