#!/usr/bin/env python3
"""
Wave 14 Hardening and Edge Case Test Suite.

Provides targeted tests for uncovered branches and edge cases in:
  - DefFileGenerator/def_gen.py
  - DefFileGenerator/extractor.py
  - generate_webdyn_def.py
"""

import os
import sys
import tempfile
import unittest
from unittest.mock import patch

import generate_webdyn_def
from DefFileGenerator.def_gen import (
    Generator,
    peek_generator,
)
from DefFileGenerator.extractor import Extractor


class TestWave14DefGen(unittest.TestCase):
    """Targeted edge-case tests for Generator and utilities in def_gen.py."""

    def test_peek_generator_none_and_empty(self):
        """Test peek_generator with None and empty generator input."""
        has_data, it = peek_generator(None)
        self.assertFalse(has_data)
        self.assertEqual(list(it), [])

        def empty_gen():
            return
            yield

        has_data, it = peek_generator(empty_gen())
        self.assertFalse(has_data)
        self.assertEqual(list(it), [])

    def test_sanitize_csv_field_edge_cases(self):
        """Test sanitize_csv_field with non-printable/surrogate chars resulting in empty or formula triggers."""
        # Non-printable surrogate characters stripped to empty
        val_surrogate = "\ud800\udfff"
        sanitized = Generator.sanitize_csv_field(val_surrogate)
        self.assertEqual(sanitized, "")

        # Fullwidth formula triggers
        val_fullwidth = "\uff1d1+1"
        sanitized_fw = Generator.sanitize_csv_field(val_fullwidth)
        self.assertTrue(sanitized_fw.startswith("'"))

        # Leading prefix trigger whitespace with newline replacement
        val_prefix_nl = "\t=SUM(1,2)\r\n"
        sanitized_pnl = Generator.sanitize_csv_field(val_prefix_nl)
        self.assertTrue(sanitized_pnl.startswith("'\t="))
        self.assertNotIn("\n", sanitized_pnl)

    def test_normalize_type_s_prefix(self):
        """Test normalize_type with signed compact types starting with 's'."""
        self.assertEqual(Generator.normalize_type("s16"), "I16")
        self.assertEqual(Generator.normalize_type("s32"), "I32")
        self.assertEqual(Generator.normalize_type("s64"), "I64")

    def test_normalize_address_val_exception_handling(self):
        """Test normalize_address_val when input string conversion or parsing throws an exception."""

        class BadObject:
            def __str__(self):
                raise RuntimeError("Bad string conversion")

        self.assertEqual(Generator.normalize_address_val(BadObject()), "")

    def test_calculate_coefficients_extreme_floats(self):
        """Test _calculate_coefficients with NaN, Inf, and extreme exponents."""
        # NaN
        a, b = Generator._calculate_coefficients(float("nan"), float("nan"), "0")
        self.assertEqual(a, "1.000000")
        self.assertEqual(b, "0.000000")

        # Inf
        a, b = Generator._calculate_coefficients(float("inf"), -float("inf"), "0")
        self.assertEqual(a, "1.000000")
        self.assertEqual(b, "0.000000")

        # Scale factor out of bounds > 100
        a, b = Generator._calculate_coefficients("2.5", "10.0", "150")
        self.assertEqual(a, "2.500000")
        self.assertEqual(b, "10.000000")

    def test_apply_address_offset_negative_logging(self):
        """Test apply_address_offset logging warnings for negative resulting addresses."""
        with self.assertLogs(level="WARNING") as cm:
            res = Generator.apply_address_offset("5", -10, line_num=42, name="Voltage")
            self.assertEqual(res, "-5")
            self.assertTrue(any("Line 42:" in msg and "Voltage" in msg for msg in cm.output))

        with self.assertLogs(level="WARNING") as cm:
            res = Generator.apply_address_offset("5", -10, line_num=None, name="Current")
            self.assertEqual(res, "-5")
            self.assertTrue(any("Current" in msg and "Line" not in msg for msg in cm.output))

    def test_check_address_overlap_bits_none(self):
        """Test _check_address_overlap with BITS type when current_bits bit_slice is None."""
        gen = Generator()
        usage = {}
        warned = set()
        # Invalid address for BITS produces None for _bit_slice
        has_overlap = gen._check_address_overlap("3", "40001", "BITS", "TestBits", 1, usage, warned)
        self.assertFalse(has_overlap)


class TestWave14Extractor(unittest.TestCase):
    """Targeted edge-case tests for Extractor methods in extractor.py."""

    def test_map_and_clean_empty_and_no_address(self):
        """Test map_and_clean with empty tables or tables lacking Address column."""
        ext = Extractor()
        # Empty table stream
        res = list(ext.map_and_clean([]))
        self.assertEqual(res, [])

        # Table with no Address column
        table_no_addr = [[{"Name": "Var1", "Val": "100"}]]
        res_no_addr = list(ext.map_and_clean(table_no_addr))
        self.assertEqual(res_no_addr, [])

    def test_extract_from_pdf_nonexistent_and_corrupt(self):
        """Test extract_from_pdf gracefully handles missing or non-PDF files."""
        ext = Extractor()
        # Nonexistent file
        gen = ext.extract_from_pdf("nonexistent_file.pdf")
        self.assertEqual(list(gen), [])

        # Corrupt file
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tf:
            tf.write(b"NOT_A_PDF_FILE")
            tf_path = tf.name

        try:
            gen = ext.extract_from_pdf(tf_path)
            tables = list(gen)
            # Consuming table generators should handle error safely
            for t in tables:
                list(t)
        finally:
            if os.path.exists(tf_path):
                os.remove(tf_path)

    def test_extract_from_xml_nonexistent_and_malformed(self):
        """Test extract_from_xml handles nonexistent and malformed XML files."""
        ext = Extractor()
        # Nonexistent file
        gen = ext.extract_from_xml("nonexistent_file.xml")
        tables = list(gen)
        self.assertEqual(len(tables), 1)
        self.assertEqual(list(tables[0]), [])

        # Malformed XML
        with tempfile.NamedTemporaryFile(
            suffix=".xml", mode="w", delete=False, encoding="utf-8"
        ) as tf:
            tf.write("<root><unclosed>tag</root>")
            tf_path = tf.name

        try:
            gen = ext.extract_from_xml(tf_path)
            tables = list(gen)
            for t in tables:
                list(t)
        finally:
            if os.path.exists(tf_path):
                os.remove(tf_path)


class TestWave14GenerateWebdynDef(unittest.TestCase):
    """Targeted edge-case tests for generate_webdyn_def.py CLI and entrypoints."""

    def test_generate_webdyn_def_cli_demo_mode(self):
        """Test generate_webdyn_def.py CLI main function when run in demo mode (fewer than 5 args)."""
        demo_in = "sample_register_map.csv"
        demo_out = "sample_output_definition.csv"
        try:
            with patch.object(sys, "argv", ["generate_webdyn_def.py"]):
                with self.assertRaises(SystemExit) as cm:
                    generate_webdyn_def.main()
                self.assertEqual(cm.exception.code, 0)
        finally:
            if os.path.exists(demo_in):
                os.remove(demo_in)
            if os.path.exists(demo_out):
                os.remove(demo_out)


if __name__ == "__main__":
    unittest.main()
