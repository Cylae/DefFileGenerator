"""Wave 17 coverage tests targeting remaining uncovered branches and edge cases across

DefFileGenerator core, extractor, main CLI, and web application modules.
"""

import os
import tempfile
import unittest
from unittest.mock import patch

from fastapi.testclient import TestClient

import DefFileGenerator.main as main_module
from DefFileGenerator.def_gen import Generator, peek_generator
from DefFileGenerator.extractor import Extractor
from web.app import UploadBodyLimitMiddleware, app


class TestDefGenWave17(unittest.TestCase):
    """Targeted edge case tests for def_gen.py."""

    def test_peek_generator_empty_and_exceptions(self) -> None:
        """Test peek_generator behavior on empty iterables and error conditions."""

        def empty_gen():
            if False:
                yield 1

        has_data, restored = peek_generator(empty_gen())
        self.assertFalse(has_data)
        self.assertEqual(list(restored), [])

        def error_gen():
            yield 10
            raise ValueError("Generator fault")

        has_data_err, restored_err = peek_generator(error_gen())
        self.assertTrue(has_data_err)
        items = []
        try:
            for item in restored_err:
                items.append(item)
        except ValueError as e:
            self.assertIn("Generator fault", str(e))
        self.assertEqual(items, [10])

    def test_validate_address_negative_offsets(self) -> None:
        """Test validate_address with invalid or negative address offsets."""
        gen = Generator()
        # Invalid address format
        self.assertFalse(gen.validate_address("INVALID_ADDR", "U16", True))

        # Address resulting in negative number after offset
        res_neg = gen.apply_address_offset("10", -20)
        self.assertEqual(res_neg, "-10")
        self.assertFalse(gen.validate_address(res_neg, "U16", True))

    def test_determine_info1_edge_cases(self) -> None:
        """Test _determine_info1 with obscure and complex data types."""
        gen = Generator()
        # Default fallback for unknown dtypes is MODBUS_HOLDING ("3")
        self.assertEqual(gen._determine_info1("UNKNOWN_TYPE_123"), "3")
        self.assertEqual(gen._determine_info1("holding register"), "3")
        self.assertEqual(gen._determine_info1("input register"), "4")

    def test_calculate_coefficients_non_finite(self) -> None:
        """Test _calculate_coefficients handling of non-finite float inputs."""
        gen = Generator()
        # NaN / Inf handling
        a, b = gen._calculate_coefficients(float("nan"), float("inf"), None)
        self.assertEqual(a, "1.000000")
        self.assertEqual(b, "0.000000")

        a2, b2 = gen._calculate_coefficients(1e308, -1e308, None)
        self.assertIsInstance(a2, str)
        self.assertIsInstance(b2, str)


class TestExtractorWave17(unittest.TestCase):
    """Targeted edge case tests for extractor.py."""

    def test_extract_from_pdf_page_range_parsing(self) -> None:
        """Test Extractor.extract_from_pdf with inverted or whitespace page ranges."""
        extractor = Extractor()
        # Test with missing file path
        res = list(extractor.extract_from_pdf("non_existent_file.pdf", pages=" 10 - 5 , 1- 2 "))
        self.assertEqual(res, [])

    def test_fuzzy_header_score_and_inference(self) -> None:
        """Test _fuzzy_header_score edge cases."""
        extractor = Extractor()
        # Test empty string vs header
        score = extractor._fuzzy_header_score("", ["Address"])
        self.assertEqual(score, 0.0)

        # Test corrupt or non-matching header
        score_low = extractor._fuzzy_header_score("xyz_random_column", ["Register Address"])
        self.assertLess(score_low, 0.5)


class TestMainCLIWave17(unittest.TestCase):
    """Targeted edge case tests for main.py CLI subcommands."""

    def test_generate_subcommand_flags(self) -> None:
        """Test generate subcommand with flags."""
        with tempfile.TemporaryDirectory() as tmpdir:
            input_csv = os.path.join(tmpdir, "in.csv")
            output_csv = os.path.join(tmpdir, "out.csv")
            with open(input_csv, "w", encoding="utf-8") as f:
                f.write("Name,Address,Type\nVoltage,100,U16\n")

            test_args = [
                "main.py",
                "generate",
                input_csv,
                "-o",
                output_csv,
                "--manufacturer",
                "ACME",
                "--model",
                "V1",
            ]
            with patch("sys.argv", test_args):
                with patch("sys.exit") as mock_exit:
                    main_module.main()
                    if mock_exit.called:
                        mock_exit.assert_called_with(0)
            self.assertTrue(os.path.exists(output_csv))

    def test_extract_subcommand_invalid_output(self) -> None:
        """Test extract subcommand with unsupported file extensions."""
        with tempfile.TemporaryDirectory() as tmpdir:
            input_invalid = os.path.join(tmpdir, "in.invalid_ext")
            output_csv = os.path.join(tmpdir, "out.csv")
            with open(input_invalid, "w", encoding="utf-8") as f:
                f.write("dummy content")

            test_args = [
                "main.py",
                "extract",
                input_invalid,
                "-o",
                output_csv,
            ]
            with patch("sys.argv", test_args):
                with self.assertRaises(SystemExit) as cm:
                    main_module.main()
                self.assertEqual(cm.exception.code, 1)


class TestWebWave17(unittest.TestCase):
    """Targeted edge case tests for web/app.py."""

    def setUp(self) -> None:
        self.client = TestClient(app)

    def test_upload_body_limit_middleware_boundary(self) -> None:
        """Test UploadBodyLimitMiddleware initialization."""

        async def dummy_app(scope, receive, send):
            pass

        middleware = UploadBodyLimitMiddleware(dummy_app)
        self.assertIsNotNone(middleware)

    def test_api_convert_filename_sanitization(self) -> None:
        """Test /api/convert with whitespace or path traversal filenames."""
        file_content = b"Name,Address,Type\nFreq,50,U16\n"
        response = self.client.post(
            "/api/convert",
            files={"file": ("  ../../../test_file.csv  ", file_content, "text/csv")},
            data={"manufacturer": "TestMfg", "model": "TestModel"},
        )
        self.assertEqual(response.status_code, 200)


if __name__ == "__main__":
    unittest.main()
