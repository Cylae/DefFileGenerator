#!/usr/bin/env python3
"""
Wave 16 Coverage & Edge-Case Hardening Unit Tests.

Tests deep edge cases, error fallbacks, and boundary conditions across:
- DefFileGenerator.def_gen
- DefFileGenerator.extractor
- DefFileGenerator.main
- web.app
"""

import os
import tempfile
import unittest
from unittest.mock import patch

from fastapi.testclient import TestClient

from DefFileGenerator.def_gen import (
    Generator,
    peek_generator,
)
from DefFileGenerator.extractor import ExtractionError, Extractor
from DefFileGenerator.main import main as cli_main
from web.app import app


class TestDefGenWave16(unittest.TestCase):
    """Targeted tests for def_gen.py edge cases."""

    def test_validate_address_exceptions(self) -> None:
        # Invalid object or string triggering ValueError / IndexError in parts
        self.assertFalse(Generator.validate_address("30001_abc", "STRING", strict=True))
        self.assertFalse(Generator.validate_address("30001_0_def", "BITS", strict=True))

    def test_get_register_count_exceptions(self) -> None:
        # Malformed STR or RAW address string that causes ValueError / IndexError
        self.assertEqual(Generator.get_register_count("STRING", "30001_abc"), 0)
        self.assertEqual(Generator.get_register_count("STR_INVALID", "30001"), 0)

    def test_apply_address_offset_exceptions(self) -> None:
        self.assertEqual(Generator.apply_address_offset(None, 10), "")
        # Address part not parseable as int or hex
        self.assertEqual(Generator.apply_address_offset("xyz_qrs", 10), "xyz_qrs")

    def test_determine_info1_variations(self) -> None:
        gen = Generator()
        self.assertEqual(gen._determine_info1("unknown_type", line_num=10), "3")
        self.assertEqual(gen._determine_info1(None), "3")  # type: ignore[arg-type]
        self.assertEqual(gen._determine_info1(""), "3")

    def test_run_generator_and_template_edge_cases(self) -> None:
        from DefFileGenerator.def_gen import GeneratorConfig, run_generator

        # Same input and output path failure
        with tempfile.NamedTemporaryFile("w", suffix=".csv", delete=False) as f:
            f.write("test")
            tmp_path = f.name

        try:
            cfg = GeneratorConfig(input_file=tmp_path, output=tmp_path)
            self.assertFalse(run_generator(cfg))

            # Non-existent input file failure
            cfg_missing = GeneratorConfig(
                input_file="/nonexistent/path/to/file.csv", output="out.csv"
            )
            self.assertFalse(run_generator(cfg_missing))
        finally:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)

    def test_peek_generator(self) -> None:
        has_data, it = peek_generator(None)
        self.assertFalse(has_data)
        self.assertEqual(list(it), [])

        has_data, it = peek_generator(iter([]))
        self.assertFalse(has_data)
        self.assertEqual(list(it), [])

        has_data, it = peek_generator(iter([1, 2, 3]))
        self.assertTrue(has_data)
        self.assertEqual(list(it), [1, 2, 3])

    def test_sanitize_csv_field_edge_cases(self) -> None:
        self.assertEqual(Generator.sanitize_csv_field(None), "")
        self.assertEqual(Generator.sanitize_csv_field(123), "123")
        self.assertEqual(Generator.sanitize_csv_field(45.67), "45.67")
        self.assertEqual(Generator.sanitize_csv_field(""), "")

        # Pure surrogate/control chars that get stripped to empty
        self.assertEqual(Generator.sanitize_csv_field("\ud800\udfff"), "")

        # Formula trigger after whitespace
        self.assertEqual(Generator.sanitize_csv_field("\t=SUM(1,2)"), "'\t=SUM(1,2)")
        self.assertEqual(Generator.sanitize_csv_field("\uff1dSUM(1,2)"), "'\uff1dSUM(1,2)")

        # Formula chars with inf/nan
        self.assertEqual(Generator.sanitize_csv_field("=inf"), "'=inf")
        self.assertEqual(Generator.sanitize_csv_field("+nan"), "'+nan")

        # Valid finite float
        self.assertEqual(Generator.sanitize_csv_field("-12.34"), "-12.34")
        self.assertEqual(Generator.sanitize_csv_field(" -12.34"), "' -12.34")

    def test_normalize_type_variations(self) -> None:
        self.assertEqual(Generator.normalize_type(None), "U16")
        self.assertEqual(Generator.normalize_type(""), "U16")
        self.assertEqual(Generator.normalize_type("str"), "STRING")
        self.assertEqual(Generator.normalize_type("string"), "STRING")
        self.assertEqual(Generator.normalize_type("u16"), "U16")
        self.assertEqual(Generator.normalize_type("s16"), "I16")
        self.assertEqual(Generator.normalize_type("uint 32"), "U32")
        self.assertEqual(Generator.normalize_type("float32 swap"), "F32_WB")
        self.assertEqual(Generator.normalize_type("bitfield16"), "U16")
        self.assertEqual(Generator.normalize_type("str20"), "STR20")
        self.assertEqual(Generator.normalize_type("string 10"), "STR10")

    def test_normalize_address_val_hex_and_ranges(self) -> None:
        self.assertEqual(Generator.normalize_address_val(100), "100")
        self.assertEqual(Generator.normalize_address_val(None), "None")
        self.assertEqual(Generator.normalize_address_val(""), "")
        self.assertEqual(Generator.normalize_address_val("40001"), "40001")
        self.assertEqual(Generator.normalize_address_val("0x1F40"), "8000")
        self.assertEqual(Generator.normalize_address_val("1F40h"), "8000")
        self.assertEqual(Generator.normalize_address_val("30,001"), "30001")
        self.assertEqual(Generator.normalize_address_val("30001..30005"), "30001")
        self.assertEqual(Generator.normalize_address_val("30001 - 30005"), "30001")
        self.assertEqual(Generator.normalize_address_val("-0x10"), "-16")
        self.assertEqual(Generator.normalize_address_val("-10h"), "-16")

    def test_validate_address_strict_vs_nonstrict(self) -> None:
        # Out of bounds address 70000
        self.assertFalse(Generator.validate_address("70000", "U16", strict=True))
        self.assertTrue(Generator.validate_address("70000", "U16", strict=False))

        # Invalid format
        self.assertFalse(Generator.validate_address("abc_def_ghi", "BITS", strict=True))
        self.assertFalse(Generator.validate_address("30001_-5", "STRING", strict=True))
        self.assertFalse(Generator.validate_address("30001_1_20", "BITS", strict=True))

    def test_calculate_coefficients_extremes(self) -> None:
        coef_a, coef_b = Generator._calculate_coefficients("1.0", "0.0", "500")
        self.assertEqual(coef_a, "1.000000")
        self.assertEqual(coef_b, "0.000000")

        coef_a, coef_b = Generator._calculate_coefficients("inf", "nan", "0")
        self.assertEqual(coef_a, "1.000000")
        self.assertEqual(coef_b, "0.000000")

        coef_a, coef_b = Generator._calculate_coefficients("0.5", "-10.5", "-1")
        self.assertEqual(coef_a, "0.050000")
        self.assertEqual(coef_b, "-10.500000")

    def test_write_output_csv_validation_failure_preserves_target(self) -> None:
        gen = Generator()
        with tempfile.TemporaryDirectory() as tmpdir:
            out_file = os.path.join(tmpdir, "output.csv")
            with open(out_file, "w") as f:
                f.write("original content")

            # Try writing zero registers with validation_strict=True
            success = gen.write_output_csv(out_file, [], validation_strict=True)
            self.assertFalse(success)
            with open(out_file) as f:
                self.assertEqual(f.read(), "original content")


class TestExtractorWave16(unittest.TestCase):
    """Targeted tests for extractor.py edge cases."""

    def test_fuzzy_header_score(self) -> None:
        # Test reverse string fuzzy score
        score = Extractor._fuzzy_header_score("s s e r d d a", ["Address"])
        self.assertGreaterEqual(score, 0.75)
        self.assertTrue(Extractor._fuzzy_header_matches("s s e r d d a", ["Address"]))

        # Short strings should yield 0.0
        self.assertEqual(Extractor._fuzzy_header_score("abc", ["Address"]), 0.0)

    def test_infer_table_columns(self) -> None:
        self.assertIsNone(Extractor._infer_table_columns([]))

        # Table with small numeric row indices that lack clear address preference
        table_ambiguous = [
            ["1", "2", "3"],
            ["4", "5", "6"],
        ]
        self.assertIsNone(Extractor._infer_table_columns(table_ambiguous))

        # Table with clear address, type, and unit columns
        table_clear = [
            ["30001", "U16", "V", "RO"],
            ["30002", "U32", "A", "RW"],
        ]
        inferred = Extractor._infer_table_columns(table_clear)
        self.assertIsNotNone(inferred)
        if inferred:
            self.assertIn("Address", inferred)
            self.assertIn("Type", inferred)

    def test_extract_from_csv_nul_bytes(self) -> None:
        extractor = Extractor()
        with tempfile.NamedTemporaryFile("wb", suffix=".csv", delete=False) as f:
            f.write(b"Address,Name\x00,Type\n30001,Test,U16")
            tmp_path = f.name

        try:
            tables = list(extractor.extract_from_csv(tmp_path))
            self.assertEqual(len(tables), 1)
            with self.assertRaises(ExtractionError):
                list(tables[0])
        finally:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)

    def test_extract_from_xml_deduplication(self) -> None:
        extractor = Extractor()
        xml_content = """<?xml version="1.0"?>
        <registers>
            <item>
                <address>30001</address>
                <name>Voltage</name>
                <type>U16</type>
            </item>
            <item>
                <address>30001</address>
                <name>Voltage</name>
                <type>U16</type>
            </item>
        </registers>
        """
        with tempfile.NamedTemporaryFile("w", suffix=".xml", delete=False, encoding="utf-8") as f:
            f.write(xml_content)
            tmp_path = f.name

        try:
            tables = list(extractor.extract_from_xml(tmp_path))
            rows = list(tables[0])
            self.assertEqual(len(rows), 1)  # Deduplicated duplicate row
            self.assertEqual(rows[0]["address"], "30001")
        finally:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)

    def test_map_and_clean_gain_inversion_and_byte_length(self) -> None:
        extractor = Extractor()
        raw_tables = [
            [
                {
                    "Register": "30001",
                    "Description": "Active Power",
                    "Data Type": "STRING",
                    "Qty": "10",
                    "Gain": "100",
                }
            ]
        ]
        mapped = list(extractor.map_and_clean(raw_tables, address_offset=0))
        self.assertEqual(len(mapped), 1)
        row = mapped[0]
        self.assertEqual(row["Address"], "30001_20")  # 10 words * 2 = 20 bytes
        self.assertEqual(row["Factor"], "0.01")  # 1 / 100 = 0.01


class TestWebAPIWave16(unittest.TestCase):
    """Targeted tests for web/app.py edge cases."""

    def setUp(self) -> None:
        self.client = TestClient(app)

    def test_health_check(self) -> None:
        response = self.client.get("/api/health")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json(), {"status": "ok", "version": "0.2.1"})

    def test_validate_unsupported_extension(self) -> None:
        response = self.client.post(
            "/api/validate",
            files={"file": ("test.pdf", b"pdf content", "application/pdf")},
        )
        self.assertEqual(response.status_code, 400)
        self.assertIn("Unsupported file format", response.json()["detail"])

    def test_convert_unsupported_extension(self) -> None:
        response = self.client.post(
            "/api/convert",
            files={"file": ("test.exe", b"exe content", "application/octet-stream")},
        )
        self.assertEqual(response.status_code, 400)
        self.assertIn("Unsupported file format", response.json()["detail"])

    def test_convert_no_readable_registers(self) -> None:
        csv_bytes = b"ColA,ColB\nValA,ValB\n"
        response = self.client.post(
            "/api/convert",
            files={"file": ("empty.csv", csv_bytes, "text/csv")},
        )
        self.assertEqual(response.status_code, 400)
        self.assertIn("No valid registers mapped", response.json()["detail"])


class TestMainCLIWave16(unittest.TestCase):
    """Targeted tests for main.py CLI subcommands."""

    def test_cli_version_flag(self) -> None:
        with patch("sys.argv", ["deffilegen", "--version"]):
            with self.assertRaises(SystemExit) as cm:
                cli_main()
            self.assertEqual(cm.exception.code, 0)

    def test_cli_validate_subcommand(self) -> None:
        with tempfile.NamedTemporaryFile("w", suffix=".csv", delete=False) as f:
            f.write("modbusRTU;Inverter;Mfg;Model;;;;;;;\n1;3;40001;U16;;P;p;1.0;0.0;W;4\n")
            tmp_path = f.name

        try:
            with patch("sys.argv", ["deffilegen", "validate", tmp_path]):
                # Successful validation exits cleanly without raising SystemExit or with 0
                try:
                    cli_main()
                except SystemExit as e:
                    self.assertEqual(e.code, 0)
        finally:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)


if __name__ == "__main__":
    unittest.main()
