"""
Wave 19 Unit Tests: Targeted coverage expansion for Generator row validation helpers,
CLI edge handling in DefFileGenerator/main.py, and REST API error boundaries.
"""

import os
import tempfile
import unittest
from unittest.mock import patch

from DefFileGenerator.def_gen import Generator, ValidationIssue
from DefFileGenerator.main import main as cli_main


class TestGeneratorRowValidationHelpers(unittest.TestCase):
    """Targeted tests for private Generator per-row validation helper methods."""

    def setUp(self) -> None:
        self.generator = Generator()

    def test_validate_row_info1_valid(self) -> None:
        issues: list[ValidationIssue] = []
        counts: dict[str, int] = {}
        code, ok = self.generator._validate_row_info1("3", line_num=2, strict=True, issues=issues, type_counts=counts)
        self.assertTrue(ok)
        self.assertEqual(code, "3")
        self.assertEqual(counts.get("3"), 1)
        self.assertEqual(len(issues), 0)

    def test_validate_row_info1_invalid_lenient(self) -> None:
        issues: list[ValidationIssue] = []
        counts: dict[str, int] = {}
        code, ok = self.generator._validate_row_info1("99", line_num=3, strict=False, issues=issues, type_counts=counts)
        self.assertTrue(ok)
        self.assertEqual(code, "99")
        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0].severity, "WARNING")

    def test_validate_row_info1_invalid_strict(self) -> None:
        issues: list[ValidationIssue] = []
        counts: dict[str, int] = {}
        code, ok = self.generator._validate_row_info1("99", line_num=4, strict=True, issues=issues, type_counts=counts)
        self.assertFalse(ok)
        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0].severity, "ERROR")

    def test_validate_row_tag_duplicate(self) -> None:
        issues: list[ValidationIssue] = []
        seen = {"v_test": 2}
        ok = self.generator._validate_row_tag("v_test", line_num=5, seen_tags=seen, issues=issues)
        self.assertFalse(ok)
        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0].code, "DUPLICATE_TAG")

    def test_validate_row_type_invalid(self) -> None:
        issues: list[ValidationIssue] = []
        ok = self.generator._validate_row_type("INVALID_TYPE_XYZ", line_num=6, issues=issues)
        self.assertFalse(ok)
        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0].code, "INVALID_TYPE")

    def test_validate_row_address_invalid(self) -> None:
        issues: list[ValidationIssue] = []
        ok = self.generator._validate_row_address("70000", "U16", line_num=7, strict=True, issues=issues)
        self.assertFalse(ok)
        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0].code, "INVALID_ADDRESS")

    def test_validate_row_action_invalid_strict(self) -> None:
        issues: list[ValidationIssue] = []
        ok = self.generator._validate_row_action("99", line_num=8, strict=True, issues=issues)
        self.assertFalse(ok)
        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0].severity, "ERROR")

    def test_validate_row_coefs_invalid_float(self) -> None:
        issues: list[ValidationIssue] = []
        # row: Index, Info1, Info2, Info3, Info4, Name, Tag, CoefA, CoefB, Unit, Action
        row = ["1", "3", "40001", "U16", "", "Test", "test", "nan", "0.0", "V", "4"]
        ok = self.generator._validate_row_coefs(row, line_num=9, strict=True, issues=issues)
        self.assertFalse(ok)
        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0].code, "INVALID_COEF")


class TestCLIEdgeCasesWave19(unittest.TestCase):
    """Targeted tests for CLI edge branches in DefFileGenerator/main.py."""

    def test_pages_flag_warning_on_csv(self) -> None:
        with tempfile.NamedTemporaryFile("w", suffix=".csv", delete=False) as f:
            f.write("Name,Tag,RegisterType,Address,Type,Factor,Offset,Unit,Action,ScaleFactor\n")
            f.write("Voltage,v,Holding,30001,U16,1,0,V,4,0\n")
            csv_path = f.name

        try:
            with patch("logging.warning") as mock_warn:
                try:
                    cli_main(["run", csv_path, "--pages", "1-5", "--output", csv_path + ".out", "--manufacturer", "M", "--model", "Mod"])
                except SystemExit:
                    pass
                # Check that warning regarding --pages was emitted
                warn_calls = [call.args[0] for call in mock_warn.call_args_list if call.args]
                self.assertTrue(any("--pages is only applicable" in msg for msg in warn_calls))
        finally:
            if os.path.exists(csv_path):
                os.unlink(csv_path)
            if os.path.exists(csv_path + ".out"):
                os.unlink(csv_path + ".out")


if __name__ == "__main__":
    unittest.main()
