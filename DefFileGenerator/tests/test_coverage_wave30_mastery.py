"""
Wave 30 Mastery Test Suite.
Targeted unit tests for edge branches and fallback paths in mars_generator, def_gen, main, extractor, and web/app.
"""

import tempfile
import unittest
from pathlib import Path

from DefFileGenerator.def_gen import Generator
from DefFileGenerator.mars_generator import (
    determine_action,
    normalize_type,
    normalize_unit,
    write_definition_file,
)


class TestMarsGeneratorUnits(unittest.TestCase):
    def test_normalize_unit_branches(self) -> None:
        self.assertEqual(normalize_unit(None), ("", 1.0))
        self.assertEqual(normalize_unit("-"), ("", 1.0))
        self.assertEqual(normalize_unit("/"), ("", 1.0))
        self.assertEqual(normalize_unit("0.01Hz"), ("Hz", 0.01))
        self.assertEqual(normalize_unit("°C"), ("°C", 1.0))
        self.assertEqual(normalize_unit("degC"), ("°C", 1.0))
        self.assertEqual(normalize_unit("ohm"), ("Ω", 1.0))
        self.assertEqual(normalize_unit("kvar"), ("kvar", 1.0))
        self.assertEqual(normalize_unit("var"), ("Var", 1.0))
        self.assertEqual(normalize_unit("kW"), ("kW", 1.0))

    def test_normalize_type_branches(self) -> None:
        self.assertEqual(normalize_type(None), "U16")
        self.assertEqual(normalize_type(""), "U16")
        self.assertEqual(normalize_type("Float32"), "F32")
        self.assertEqual(normalize_type("u32"), "U32")
        self.assertEqual(normalize_type("i32"), "I32")
        self.assertEqual(normalize_type("int32"), "I32")
        self.assertEqual(normalize_type("i16"), "I16")
        self.assertEqual(normalize_type("int16"), "I16")
        self.assertEqual(normalize_type("uint16"), "U16")
        self.assertEqual(normalize_type("unknown_custom_type"), "U16")

    def test_determine_action_branches(self) -> None:
        self.assertEqual(determine_action("Reserve Point", "4", is_reserve=True), "0")
        self.assertEqual(determine_action("Active Power Mode Set", "3", is_reserve=False), "1")
        self.assertEqual(determine_action("System Alarm Fault", "4", is_reserve=False), "9")
        self.assertEqual(determine_action("Normal Register", "4", is_reserve=False, default_action="4"), "4")

    def test_write_definition_file_and_validation(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            out_file = Path(tmpdir) / "test_def.csv"
            registers = [
                {
                    "Info1": "4",
                    "Info2": "0",
                    "Info3": "U16",
                    "Name": "Status Register",
                    "Tag": "",
                    "CoefA": 1.0,
                    "CoefB": 0.0,
                    "Unit": "",
                    "Action": "9",
                }
            ]
            valid, report = write_definition_file(
                out_file, "modbusTCP", "Inverter", "TestMfg", "TestModel", registers
            )
            self.assertTrue(valid)
            self.assertTrue(out_file.exists())
            self.assertTrue(report.is_valid)


class TestMainAndCliEdgePaths(unittest.TestCase):
    def test_main_cli_help_or_version(self) -> None:
        from DefFileGenerator.main import main

        with self.assertRaises(SystemExit) as ctx:
            main(["--help"])
        self.assertEqual(ctx.exception.code, 0)

    def test_main_cli_extract_invalid_args(self) -> None:
        from DefFileGenerator.main import main

        with self.assertRaises(SystemExit) as ctx:
            main(["extract", "--non-existent-flag"])
        self.assertNotEqual(ctx.exception.code, 0)


class TestDefGenExtraEdge(unittest.TestCase):
    def test_generator_validate_csv_detailed_missing_file(self) -> None:
        gen = Generator()
        report = gen.validate_csv_detailed("non_existent_file_path_12345.csv")
        self.assertFalse(report.is_valid)
        self.assertTrue(len(report.issues) > 0)


class TestWebEdgeCoverage(unittest.TestCase):
    def test_web_app_index_endpoint(self) -> None:
        from fastapi.testclient import TestClient

        from web.app import app

        client = TestClient(app)
        res = client.get("/")
        self.assertEqual(res.status_code, 200)


if __name__ == "__main__":
    unittest.main()
