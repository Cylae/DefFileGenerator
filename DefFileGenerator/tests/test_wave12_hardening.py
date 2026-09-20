#!/usr/bin/env python3
"""
Wave 12 Hardening Test Suite for DefFileGenerator.

Provides targeted unit tests for:
- doc_to_webdyn.py CLI edge cases (invalid mapping files, invalid page formats, unsupported extensions).
- generate_webdyn_def.py error handling (missing arguments, non-existent files, unparseable input).
- web/app.py REST API edge cases (filename sanitization, unsupported formats, empty uploads).
"""

import os
import tempfile
import unittest
from unittest.mock import patch

from fastapi.testclient import TestClient

import doc_to_webdyn
from generate_webdyn_def import generate_webdyn_definition
from web.app import app


class TestWave12Hardening(unittest.TestCase):
    """Wave 12 test suite for CLI and web API edge cases."""

    def test_doc_to_webdyn_invalid_mapping_file(self) -> None:
        """Verify doc_to_webdyn exits gracefully on a non-existent mapping file."""
        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as f:
            f.write(b"Register,Name,Data Type\n40001,Power,uint16\n")
            input_path = f.name

        try:
            with patch("logging.error"):
                with self.assertRaises(SystemExit) as cm:
                    doc_to_webdyn._run_cli([input_path, "--mapping", "non_existent_mapping.json"])
                self.assertEqual(cm.exception.code, 1)
        finally:
            if os.path.exists(input_path):
                os.remove(input_path)

    def test_doc_to_webdyn_invalid_pages_format(self) -> None:
        """Verify doc_to_webdyn exits gracefully on invalid --pages format."""
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
            f.write(b"%PDF-1.4 fake pdf content")
            input_path = f.name

        try:
            with patch("logging.error"):
                with self.assertRaises(SystemExit) as cm:
                    doc_to_webdyn._run_cli([input_path, "--pages", "invalid,pages,format"])
                self.assertEqual(cm.exception.code, 1)
        finally:
            if os.path.exists(input_path):
                os.remove(input_path)

    def test_doc_to_webdyn_unsupported_extension(self) -> None:
        """Verify doc_to_webdyn exits gracefully on unsupported file extensions."""
        with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as f:
            f.write(b"some text data")
            input_path = f.name

        try:
            with patch("logging.error"):
                with self.assertRaises(SystemExit) as cm:
                    doc_to_webdyn._run_cli([input_path])
                self.assertEqual(cm.exception.code, 1)
        finally:
            if os.path.exists(input_path):
                os.remove(input_path)

    def test_generate_webdyn_def_missing_args(self) -> None:
        """Verify generate_webdyn_definition returns False when required args are missing."""
        with patch("logging.error"):
            result = generate_webdyn_definition(input_file="some_file.csv")
            self.assertFalse(result)

    def test_generate_webdyn_def_nonexistent_input(self) -> None:
        """Verify generate_webdyn_definition returns False for non-existent input files."""
        with patch("logging.error"):
            result = generate_webdyn_definition(
                input_file="non_existent_file.csv",
                output_file="out.csv",
                manufacturer="Mfg",
                model="Model",
            )
            self.assertFalse(result)

    def test_generate_webdyn_def_unparseable_content(self) -> None:
        """Verify generate_webdyn_definition returns False when no registers can be extracted."""
        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as f:
            f.write(b"Header1,Header2\n")
            input_path = f.name

        out_path = input_path + ".out.csv"
        try:
            with patch("logging.error"):
                result = generate_webdyn_definition(
                    input_file=input_path,
                    output_file=out_path,
                    manufacturer="Mfg",
                    model="Model",
                )
                self.assertFalse(result)
        finally:
            if os.path.exists(input_path):
                os.remove(input_path)
            if os.path.exists(out_path):
                os.remove(out_path)

    def test_web_app_convert_empty_filename(self) -> None:
        """Verify /api/convert returns 400 for empty or whitespace filenames."""
        client = TestClient(app)
        response = client.post(
            "/api/convert",
            files={"file": ("   ", b"data", "text/csv")},
            data={"manufacturer": "Test", "model": "TestModel"},
        )
        self.assertEqual(response.status_code, 400)

    def test_web_app_validate_unsupported_format(self) -> None:
        """Verify /api/validate returns 400 for unsupported extension files."""
        client = TestClient(app)
        response = client.post(
            "/api/validate",
            files={"file": ("test.pdf", b"data", "application/pdf")},
        )
        self.assertEqual(response.status_code, 400)


if __name__ == "__main__":
    unittest.main()
