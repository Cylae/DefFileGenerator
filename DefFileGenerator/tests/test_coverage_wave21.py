#!/usr/bin/env python3
"""
Wave 21 unit tests targeting edge cases and unreached coverage branches in:
- DefFileGenerator/def_gen.py
- DefFileGenerator/extractor.py
- DefFileGenerator/main.py
- web/app.py
"""

import csv
import os
import sys
import tempfile
import unittest
from typing import Any
from unittest.mock import patch

from fastapi.testclient import TestClient

from DefFileGenerator.def_gen import (
    Generator,
    GeneratorConfig,
    RegisterEntry,
    ValidationIssue,
    generate_template,
    peek_generator,
    run_generator,
)
from DefFileGenerator.def_gen import (
    main as def_gen_main,
)
from web.app import app


class TestDefGenWave21(unittest.TestCase):
    """Targeted tests for def_gen.py edge branches."""

    def test_peek_generator_none_and_empty(self) -> None:
        has_data, it = peek_generator(None)
        self.assertFalse(has_data)
        self.assertEqual(list(it), [])

        has_data, it = peek_generator([])
        self.assertFalse(has_data)
        self.assertEqual(list(it), [])

        has_data, it = peek_generator([1, 2, 3])
        self.assertTrue(has_data)
        self.assertEqual(list(it), [1, 2, 3])

    def test_validate_type_edge_cases(self) -> None:
        gen = Generator()
        self.assertTrue(gen.validate_type("U16"))
        self.assertTrue(gen.validate_type("BITS"))
        self.assertTrue(gen.validate_type("STR20"))
        self.assertFalse(gen.validate_type("INVALID_TYPE_XYZ"))

    def test_validate_address_invalid_bounds(self) -> None:
        gen = Generator()
        # STRING / RAW with negative length
        self.assertFalse(gen.validate_address("40001_-5", "STRING"))
        self.assertFalse(gen.validate_address("40001_0", "RAW"))
        # RAW with odd length
        self.assertFalse(gen.validate_address("40001_3", "RAW"))
        # BITS with slice exceeding 16 bits (start=10, len=10 => end=20 > 16)
        self.assertFalse(gen.validate_address("40001_10_10", "BITS"))
        # Address out of range with strict=False
        self.assertTrue(gen.validate_address("70000", "U16", strict=False))
        self.assertFalse(gen.validate_address("70000", "U16", strict=True))

    def test_apply_address_offset_logging(self) -> None:
        gen = Generator()
        with self.assertLogs(level="WARNING") as cm:
            res = gen.apply_address_offset("10", -20, line_num=5, name="test_var")
            self.assertEqual(res, "-10")
            self.assertTrue(
                any(
                    "Line 5: Address offset -20 results in negative address -10 for 'test_var'"
                    in msg
                    for msg in cm.output
                )
            )

    def test_check_address_overlap_register_entry(self) -> None:
        gen = Generator()
        entry1 = RegisterEntry(info1="3", address="40001", dtype="U16", name="var1", line_num=2)
        entry2 = RegisterEntry(info1="3", address="40001", dtype="U16", name="var2", line_num=3)
        usage: dict[str, dict[str, Any]] = {}
        warned: set[tuple[int, int]] = set()

        overlap1 = gen._check_address_overlap(entry1, address_usage=usage, warned_lines=warned)
        self.assertFalse(overlap1)

        overlap2 = gen._check_address_overlap(entry2, address_usage=usage, warned_lines=warned)
        self.assertTrue(overlap2)

    def test_validate_row_coefs_non_finite(self) -> None:
        gen = Generator()
        issues: list[ValidationIssue] = []
        # Index 7 is CoefA, Index 8 is CoefB in the definition row
        row = ["1", "3", "40001", "U16", "", "name", "tag", "nan", "0.0", "", "4"]
        valid = gen._validate_row_coefs(row, line_num=2, strict=True, issues=issues)
        self.assertFalse(valid)
        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0].code, "INVALID_COEF")

    def test_validate_csv_detailed_io_error(self) -> None:
        gen = Generator()
        report = gen.validate_csv_detailed("non_existent_file_path_12345.csv")
        self.assertFalse(report.is_valid)
        self.assertEqual(report.issues[0].code, "FILE_NOT_FOUND")

    def test_validate_csv_detailed_invalid_header(self) -> None:
        gen = Generator()
        with tempfile.NamedTemporaryFile("w", delete=False, encoding="utf-8") as tf:
            tf.write("\n")
            tf_path = tf.name

        try:
            report = gen.validate_csv_detailed(tf_path)
            self.assertFalse(report.is_valid)
            self.assertEqual(report.issues[0].code, "INVALID_HEADER")
        finally:
            if os.path.exists(tf_path):
                os.unlink(tf_path)

    def test_write_output_csv_validation_strict_empty(self) -> None:
        gen = Generator()
        with tempfile.NamedTemporaryFile("w", delete=False, encoding="utf-8") as tf:
            tf_path = tf.name

        try:
            # Processed rows empty with validation_strict=True should return False
            success = gen.write_output_csv(
                tf_path, [], config=GeneratorConfig(), validation_strict=True
            )
            self.assertFalse(success)
        finally:
            if os.path.exists(tf_path):
                os.unlink(tf_path)

    def test_generate_template_invalid_path(self) -> None:
        res = generate_template("/invalid_dir_9999/test_template.csv")
        self.assertFalse(res)

    def test_run_generator_missing_input_file(self) -> None:
        cfg = GeneratorConfig(input_file="non_existent_file_9999.csv", output="out.csv")
        res = run_generator(cfg)
        self.assertFalse(res)

    def test_run_generator_same_input_and_output(self) -> None:
        with tempfile.NamedTemporaryFile("w", delete=False) as tf:
            tf_path = tf.name

        try:
            cfg = GeneratorConfig(input_file=tf_path, output=tf_path)
            res = run_generator(cfg)
            self.assertFalse(res)
        finally:
            if os.path.exists(tf_path):
                os.unlink(tf_path)

    def test_def_gen_main_cli_execution(self) -> None:
        with tempfile.NamedTemporaryFile("w", delete=False, suffix=".csv") as tf:
            writer = csv.writer(tf)
            writer.writerow(
                [
                    "Name",
                    "Tag",
                    "RegisterType",
                    "Address",
                    "Type",
                    "Factor",
                    "Offset",
                    "Unit",
                    "Action",
                    "ScaleFactor",
                ]
            )
            writer.writerow(
                ["TestVar", "test_var", "Holding Register", "40001", "U16", "1", "0", "V", "4", "0"]
            )
            tf_path = tf.name

        with tempfile.NamedTemporaryFile("w", delete=False, suffix=".csv") as tf_out:
            out_path = tf_out.name

        try:
            test_args = [
                "def_gen.py",
                tf_path,
                "-o",
                out_path,
                "--manufacturer",
                "TestMfg",
                "--model",
                "TestMod",
            ]
            with patch.object(sys, "argv", test_args):
                def_gen_main()
            self.assertTrue(os.path.exists(out_path))
            self.assertGreater(os.path.getsize(out_path), 0)
        finally:
            if os.path.exists(tf_path):
                os.unlink(tf_path)
            if os.path.exists(out_path):
                os.unlink(out_path)


class TestWebAppWave21(unittest.TestCase):
    """Targeted tests for web/app.py edge branches."""

    def test_root_endpoint_without_index(self) -> None:
        client = TestClient(app)
        response = client.get("/")
        self.assertEqual(response.status_code, 200)


if __name__ == "__main__":
    unittest.main()
