"""
Permanent regression battery for Wave 10 codebase audit and hardening.

Covers:
1. Surplus column handling in CSV extraction (reproducing and preventing AttributeError on d_row[None] list).
2. Direct script execution of gui.py with CLI inspection flags (--version, --help).
3. build_exe.py virtual environment PyInstaller detection and delegation.
4. Pathological CSVs with leading empty fields and variable column counts.
"""

from __future__ import annotations

import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from DefFileGenerator.extractor import Extractor

REPO_ROOT = Path(__file__).resolve().parent.parent.parent


class TestWave10Hardening(unittest.TestCase):
    def setUp(self) -> None:
        self.tmpdir = tempfile.TemporaryDirectory()
        self.tmp_path = Path(self.tmpdir.name)

    def tearDown(self) -> None:
        self.tmpdir.cleanup()

    def test_csv_extraction_with_surplus_columns_and_empty_leading_cells(self) -> None:
        """
        Verify that CSV rows with empty leading values and surplus trailing columns
        (which causes csv.DictReader to assign a list to key None) are processed
        gracefully without raising AttributeError: 'list' object has no attribute 'strip'.
        """
        csv_file = self.tmp_path / "surplus_cols.csv"
        # Row 1: Headers (3 columns)
        # Row 2: Empty leading cells, then surplus extra values (e.g. 5 columns total)
        # Row 3: Normal valid row
        content = (
            "Name,Address,Type\n"
            ",,30001,Extra1,Extra2\n"
            "Voltage,30002,U16,ExtraTrailing\n"
            "Current,30003,U16\n"
        )
        csv_file.write_text(content, encoding="utf-8")

        extractor = Extractor()
        tables = list(extractor.extract_from_csv(str(csv_file)))
        self.assertEqual(len(tables), 1, "Expected exactly 1 extracted table generator")

        rows = list(tables[0])
        # Verify rows were yielded without any unhandled AttributeError
        self.assertGreaterEqual(len(rows), 2)
        # Check that None key is not leaked into yielded row dictionaries
        for r in rows:
            self.assertNotIn(
                None, r, "None key from surplus DictReader columns must be filtered out"
            )

    def test_csv_extraction_all_empty_row_with_surplus_empty_cells(self) -> None:
        """Verify entirely whitespace or blank rows with irregular delimiters don't crash."""
        csv_file = self.tmp_path / "ragged_empty.csv"
        content = "Name,Address,Type\n,,,,\n   ,   ,   ,   \nFrequency,30050,U32\n"
        csv_file.write_text(content, encoding="utf-8")

        extractor = Extractor()
        tables = list(extractor.extract_from_csv(str(csv_file)))
        rows = list(tables[0])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Name"], "Frequency")
        self.assertEqual(rows[0]["Address"], "30050")

    def test_gui_direct_execution_version_flag(self) -> None:
        """Verify that DefFileGenerator/gui.py can be invoked directly as a script with --version."""
        gui_script = REPO_ROOT / "DefFileGenerator" / "gui.py"
        res = subprocess.run(
            [sys.executable, str(gui_script), "--version"],
            capture_output=True,
            text=True,
            check=True,
        )
        self.assertIn("deffilegen-gui 0.2.1", res.stdout)

    def test_gui_direct_execution_help_flag(self) -> None:
        """Verify that DefFileGenerator/gui.py can be invoked directly as a script with --help."""
        gui_script = REPO_ROOT / "DefFileGenerator" / "gui.py"
        res = subprocess.run(
            [sys.executable, str(gui_script), "--help"],
            capture_output=True,
            text=True,
            check=True,
        )
        self.assertIn("WebdynSunPM Definition Generator", res.stdout)
        self.assertIn("--help", res.stdout)
        self.assertIn("--version", res.stdout)

    def test_build_exe_check_pyinstaller_succeeds_in_current_env(self) -> None:
        """Verify check_pyinstaller in build_exe runs without error when PyInstaller is present."""
        sys.path.insert(0, str(REPO_ROOT))
        import build_exe

        # Should execute cleanly without calling sys.exit
        build_exe.check_pyinstaller()


if __name__ == "__main__":
    unittest.main()
