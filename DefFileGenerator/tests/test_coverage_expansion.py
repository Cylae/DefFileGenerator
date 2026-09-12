import argparse
import logging
import os
import tempfile
import unittest
from unittest.mock import patch

from DefFileGenerator.extractor import Extractor, peek_generator
from DefFileGenerator.main import (
    _perform_extraction,
    extract_command,
    generate_command,
    main,
    run_command,
    setup_logging,
    validate_command,
)


class TestExtractorEdgeCases(unittest.TestCase):
    def setUp(self):
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)

    def test_peek_generator_func(self):
        # Test None
        has_data, it = peek_generator(None)
        self.assertFalse(has_data)
        self.assertEqual(list(it), [])

        # Test empty list
        has_data, it = peek_generator([])
        self.assertFalse(has_data)
        self.assertEqual(list(it), [])

        # Test non-empty list
        has_data, it = peek_generator([1, 2, 3])
        self.assertTrue(has_data)
        self.assertEqual(list(it), [1, 2, 3])

    def test_extractor_missing_openpyxl(self):
        with patch("DefFileGenerator.extractor.HAS_OPENPYXL", False):
            extractor = Extractor()
            res = list(extractor.extract_from_excel("dummy.xlsx"))
            self.assertEqual(res, [])

    def test_extractor_missing_pdfplumber(self):
        with patch("DefFileGenerator.extractor.HAS_PDFPLUMBER", False):
            extractor = Extractor()
            res = list(extractor.extract_from_pdf("dummy.pdf"))
            self.assertEqual(res, [])

    def test_extractor_missing_defusedxml(self):
        with patch("DefFileGenerator.extractor.HAS_DEFUSEDXML", False):
            extractor = Extractor()
            res = list(extractor.extract_from_xml("dummy.xml"))
            self.assertEqual(res, [])

    def test_extractor_excel_missing_file(self):
        extractor = Extractor()
        sheets = list(extractor.extract_from_excel("non_existent_file_xyz.xlsx"))
        self.assertEqual(len(sheets), 1)
        # Consuming sheet iterator yields empty
        self.assertEqual(list(sheets[0]), [])

    def test_extractor_excel_bad_zip(self):
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            f.write(b"not a zip file")
            tmp_path = f.name
        try:
            extractor = Extractor()
            sheets = list(extractor.extract_from_excel(tmp_path))
            self.assertEqual(len(sheets), 1)
            self.assertEqual(list(sheets[0]), [])
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def test_extractor_csv_file_io_error(self):
        extractor = Extractor()
        tables = list(extractor.extract_from_csv("non_existent_file_xyz.csv"))
        self.assertEqual(len(tables), 1)
        self.assertEqual(list(tables[0]), [])

    def test_extractor_csv_unicode_error(self):
        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as f:
            f.write(b"\x80\x81\x82\x83")
            tmp_path = f.name
        try:
            extractor = Extractor()
            tables = list(extractor.extract_from_csv(tmp_path))
            self.assertEqual(len(tables), 1)
            self.assertEqual(list(tables[0]), [])
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def test_extractor_xml_missing_file(self):
        extractor = Extractor()
        tables = list(extractor.extract_from_xml("non_existent_file_xyz.xml"))
        self.assertEqual(len(tables), 1)
        self.assertEqual(list(tables[0]), [])

    def test_extractor_xml_parse_error(self):
        with tempfile.NamedTemporaryFile(suffix=".xml", delete=False) as f:
            f.write(b"<invalid xml format")
            tmp_path = f.name
        try:
            extractor = Extractor()
            tables = list(extractor.extract_from_xml(tmp_path))
            self.assertEqual(len(tables), 1)
            self.assertEqual(list(tables[0]), [])
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def test_extractor_cli_main_unsupported(self):
        with patch("sys.argv", ["extractor.py", "invalid_file.unsupported_ext"]):
            with self.assertRaises(SystemExit) as cm:
                from DefFileGenerator.extractor import main as extractor_main

                extractor_main()
            self.assertEqual(cm.exception.code, 1)

    def test_extractor_cli_main_csv(self):
        with tempfile.NamedTemporaryFile(
            suffix=".csv", mode="w+", delete=False, encoding="utf-8"
        ) as f:
            f.write("Name,Address,Type\nVar1,100,U16\n")
            in_csv = f.name
        out_csv = in_csv + ".out.csv"

        try:
            with patch("sys.argv", ["extractor.py", in_csv, "-o", out_csv]):
                from DefFileGenerator.extractor import main as extractor_main

                extractor_main()
            self.assertTrue(os.path.exists(out_csv))
        finally:
            if os.path.exists(in_csv):
                os.remove(in_csv)
            if os.path.exists(out_csv):
                os.remove(out_csv)


class TestMainCliCoverage(unittest.TestCase):
    def setUp(self):
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)

    def test_setup_logging(self):
        setup_logging(verbose=True, quiet=False)
        setup_logging(verbose=False, quiet=True)
        setup_logging(verbose=False, quiet=False)

    def test_perform_extraction_missing_input_file(self):
        args = argparse.Namespace(input_file=None)
        with self.assertRaises(SystemExit) as cm:
            _perform_extraction(args)
        self.assertEqual(cm.exception.code, 1)

    def test_perform_extraction_non_existent_file(self):
        args = argparse.Namespace(input_file="non_existent_123.csv")
        with self.assertRaises(SystemExit) as cm:
            _perform_extraction(args)
        self.assertEqual(cm.exception.code, 1)

    def test_perform_extraction_unsupported_ext(self):
        with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as f:
            f.write(b"dummy")
            tmp_path = f.name
        try:
            args = argparse.Namespace(input_file=tmp_path)
            with self.assertRaises(SystemExit) as cm:
                _perform_extraction(args)
            self.assertEqual(cm.exception.code, 1)
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def test_perform_extraction_invalid_mapping_json(self):
        with (
            tempfile.NamedTemporaryFile(suffix=".csv", mode="w+", delete=False) as f_csv,
            tempfile.NamedTemporaryFile(suffix=".json", mode="w+", delete=False) as f_json,
        ):
            f_csv.write("Name,Address\nVar1,100\n")
            f_json.write("{invalid json")
            csv_path = f_csv.name
            json_path = f_json.name

        try:
            args = argparse.Namespace(input_file=csv_path, mapping=json_path)
            with self.assertRaises(SystemExit) as cm:
                _perform_extraction(args)
            self.assertEqual(cm.exception.code, 1)
        finally:
            if os.path.exists(csv_path):
                os.remove(csv_path)
            if os.path.exists(json_path):
                os.remove(json_path)

    def test_perform_extraction_empty_csv_data(self):
        with tempfile.NamedTemporaryFile(suffix=".csv", mode="w+", delete=False) as f:
            f.write("")
            tmp_path = f.name
        try:
            args = argparse.Namespace(input_file=tmp_path)
            with self.assertRaises(SystemExit) as cm:
                _perform_extraction(args)
            self.assertEqual(cm.exception.code, 1)
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def test_extract_command_existing_file_no_force(self):
        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as f:
            out_path = f.name
        try:
            args = argparse.Namespace(output=out_path, force=False)
            with self.assertRaises(SystemExit) as cm:
                extract_command(args)
            self.assertEqual(cm.exception.code, 1)
        finally:
            if os.path.exists(out_path):
                os.remove(out_path)

    def test_validate_command_missing_file(self):
        args = argparse.Namespace(input_file="non_existent_file.csv")
        with self.assertRaises(SystemExit) as cm:
            validate_command(args)
        self.assertEqual(cm.exception.code, 1)

    def test_generate_command_missing_mfg_or_model(self):
        args = argparse.Namespace(
            template=False, manufacturer=None, model="M1", input_file="in.csv"
        )
        with self.assertRaises(SystemExit) as cm:
            generate_command(args)
        self.assertEqual(cm.exception.code, 1)

    def test_generate_command_missing_input_file(self):
        args = argparse.Namespace(template=False, manufacturer="Mfg", model="M1", input_file=None)
        with self.assertRaises(SystemExit) as cm:
            generate_command(args)
        self.assertEqual(cm.exception.code, 1)

    def test_generate_command_non_existent_input_file(self):
        args = argparse.Namespace(
            template=False, manufacturer="Mfg", model="M1", input_file="missing.csv"
        )
        with self.assertRaises(SystemExit) as cm:
            generate_command(args)
        self.assertEqual(cm.exception.code, 1)

    def test_generate_command_output_exists_no_force(self):
        with (
            tempfile.NamedTemporaryFile(suffix=".csv", mode="w+", delete=False) as f_in,
            tempfile.NamedTemporaryFile(suffix=".csv", mode="w+", delete=False) as f_out,
        ):
            f_in.write("Name,Address\nVar1,100\n")
            in_path = f_in.name
            out_path = f_out.name

        try:
            args = argparse.Namespace(
                template=False,
                manufacturer="Mfg",
                model="M1",
                input_file=in_path,
                output=out_path,
                force=False,
            )
            with self.assertRaises(SystemExit) as cm:
                generate_command(args)
            self.assertEqual(cm.exception.code, 1)
        finally:
            if os.path.exists(in_path):
                os.remove(in_path)
            if os.path.exists(out_path):
                os.remove(out_path)

    def test_run_command_output_exists_no_force(self):
        with (
            tempfile.NamedTemporaryFile(suffix=".csv", mode="w+", delete=False) as f_in,
            tempfile.NamedTemporaryFile(suffix=".csv", mode="w+", delete=False) as f_out,
        ):
            f_in.write("Name,Address,Type\nVar1,100,U16\n")
            in_path = f_in.name
            out_path = f_out.name

        try:
            args = argparse.Namespace(
                template=False,
                manufacturer="Mfg",
                model="M1",
                input_file=in_path,
                output=out_path,
                force=False,
                mapping=None,
                sheet=None,
                pages=None,
                protocol="modbusRTU",
                category="Inverter",
                forced_write="",
                no_validate=False,
            )
            with self.assertRaises(SystemExit) as cm:
                run_command(args)
            self.assertEqual(cm.exception.code, 1)
        finally:
            if os.path.exists(in_path):
                os.remove(in_path)
            if os.path.exists(out_path):
                os.remove(out_path)

    def test_main_keyboard_interrupt(self):
        with patch("DefFileGenerator.main._run_cli", side_effect=KeyboardInterrupt):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 130)

    def test_main_unexpected_exception(self):
        with patch("DefFileGenerator.main._run_cli", side_effect=ValueError("Test Exception")):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)


if __name__ == "__main__":
    unittest.main()
