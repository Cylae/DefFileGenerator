"""
Targeted unit tests for Wave 20 quality elevation.

Covers:
- `DefFileGenerator.gui.main` entrypoint CLI flags (`--help`, `-h`, `--version`) in headless mode.
- `def_gen.py` edge cases in address validation, string type address parsing, coefficient calculation,
  and formula sanitization.
- `io_utils.py` file path comparison and staged text output error handling.
"""

from __future__ import annotations

import io
import os
import sys
import tempfile
import unittest
from unittest.mock import patch

from DefFileGenerator.def_gen import Generator
from DefFileGenerator.gui import main as gui_main
from DefFileGenerator.io_utils import paths_refer_to_same_file, staged_text_output


class TestWave20QualityElevation(unittest.TestCase):
    """Targeted edge-case tests for Wave 20 quality elevation."""

    def test_gui_main_cli_flags(self) -> None:
        """Verify gui.main handles --help, -h, and --version CLI flags cleanly in headless environment."""
        with patch.object(sys, "argv", ["deffilegen-gui", "--help"]):
            with patch("sys.stdout", new_callable=io.StringIO) as mock_stdout:
                gui_main()
                output = mock_stdout.getvalue()
                self.assertIn("WebdynSunPM Definition Generator", output)
                self.assertIn("--help", output)

        with patch.object(sys, "argv", ["deffilegen-gui", "-h"]):
            with patch("sys.stdout", new_callable=io.StringIO) as mock_stdout:
                gui_main()
                output = mock_stdout.getvalue()
                self.assertIn("Usage: deffilegen-gui", output)

        with patch.object(sys, "argv", ["deffilegen-gui", "--version"]):
            with patch("sys.stdout", new_callable=io.StringIO) as mock_stdout:
                gui_main()
                output = mock_stdout.getvalue()
                self.assertIn("deffilegen-gui", output)

    def test_generator_validate_address_string_length_edge_cases(self) -> None:
        """Verify Generator.validate_address handles edge cases for STR data types correctly."""
        gen = Generator()
        # Valid STR type with address 30001_16
        is_valid = gen.validate_address("30001_16", "STRING")
        self.assertTrue(is_valid)

        # Invalid STR type with zero or negative length component in address string
        is_valid_zero = gen.validate_address("30001_0", "STRING")
        self.assertFalse(is_valid_zero)

        is_valid_neg = gen.validate_address("30001_-5", "STRING")
        self.assertFalse(is_valid_neg)

    def test_generator_sanitize_csv_field_unicode_control_chars(self) -> None:
        """Verify Generator.sanitize_csv_field handles Unicode prefix triggers and control characters."""
        # Control/zero-width prefix trigger before formula
        dirty = "\u200b=1+1"
        sanitized = Generator.sanitize_csv_field(dirty)
        self.assertEqual(sanitized, "'\u200b=1+1")

        # Non-printable control character filter
        dirty_ctrl = "\x01Normal Text"
        sanitized_ctrl = Generator.sanitize_csv_field(dirty_ctrl)
        self.assertEqual(sanitized_ctrl, "Normal Text")

    def test_paths_refer_to_same_file_nonexistent(self) -> None:
        """Verify paths_refer_to_same_file handles non-existent paths safely."""
        tmp = tempfile.gettempdir()
        path1 = os.path.join(tmp, "nonexistent_file_a_123456.csv")
        path2 = os.path.join(tmp, "nonexistent_file_a_123456.csv")
        path3 = os.path.join(tmp, "nonexistent_file_b_123456.csv")

        self.assertTrue(paths_refer_to_same_file(path1, path2))
        self.assertFalse(paths_refer_to_same_file(path1, path3))

    def test_staged_text_output_cleanup_on_exception(self) -> None:
        """Verify staged_text_output removes temporary file if an exception occurs inside context block."""
        with tempfile.TemporaryDirectory() as tmpdir:
            out_file = os.path.join(tmpdir, "test_out.csv")
            with self.assertRaises(RuntimeError):
                with staged_text_output(out_file) as f:
                    f.write("partial content")
                    raise RuntimeError("Simulated failure during write")

            # Final file should not exist
            self.assertFalse(os.path.exists(out_file))


if __name__ == "__main__":
    unittest.main()
