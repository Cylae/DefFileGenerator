"""Batch 2: Deep CLI (main.py) tests – uncovered paths."""

import csv
import io
import logging
import os
import sys
import tempfile
import unittest
from unittest.mock import patch


REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))


class TestCliDeep(unittest.TestCase):

    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        self.tmpdir.cleanup()
        logging.disable(logging.NOTSET)

    def _csv_path(self, content="Address,Name,Type\n100,Var1,U16\n", name="input.csv"):
        p = os.path.join(self.tmpdir.name, name)
        with open(p, "w", encoding="utf-8") as f:
            f.write(content)
        return p

    def _def_path(self, name="def.csv"):
        p = os.path.join(self.tmpdir.name, name)
        with open(p, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(["modbusRTU","Inverter","MFG","MODEL","","","","","","",""])
            writer.writerow(["1","3","100","U16","","Var1","tag1","1.0","0.0","V","4"])
        return p

    # ------------------------------------------------------------------
    # no subcommand
    # ------------------------------------------------------------------
    def test_no_subcommand_exits_2(self):
        from DefFileGenerator.main import main
        with self.assertRaises(SystemExit) as cm:
            main([])
        self.assertEqual(cm.exception.code, 2)

    # ------------------------------------------------------------------
    # validate
    # ------------------------------------------------------------------
    def test_validate_missing_file_exits_1(self):
        from DefFileGenerator.main import main
        with self.assertRaises(SystemExit) as cm:
            main(["validate", "/no/such/file.csv"])
        self.assertEqual(cm.exception.code, 1)

    def test_validate_valid_file_succeeds(self):
        from DefFileGenerator.main import main
        p = self._def_path()
        # Should NOT raise
        main(["validate", p])

    def test_validate_lenient_flag(self):
        """--lenient should not cause a crash even if no overlaps exist."""
        from DefFileGenerator.main import main
        p = self._def_path()
        main(["validate", p, "--lenient"])

    # ------------------------------------------------------------------
    # extract
    # ------------------------------------------------------------------
    def test_extract_to_file(self):
        from DefFileGenerator.main import main
        src = self._csv_path()
        out = os.path.join(self.tmpdir.name, "out.csv")
        main(["extract", src, "-o", out])
        self.assertTrue(os.path.exists(out))

    def test_extract_missing_input_exits_1(self):
        from DefFileGenerator.main import main
        with self.assertRaises(SystemExit) as cm:
            main(["extract", "/no/such/file.csv"])
        self.assertEqual(cm.exception.code, 1)

    def test_extract_unsupported_extension_exits_1(self):
        p = os.path.join(self.tmpdir.name, "data.docx")
        with open(p, "w") as f:
            f.write("dummy")
        from DefFileGenerator.main import main
        with self.assertRaises(SystemExit) as cm:
            main(["extract", p])
        self.assertEqual(cm.exception.code, 1)

    def test_extract_guard_output_prevents_overwrite(self):
        from DefFileGenerator.main import main
        src = self._csv_path()
        out = os.path.join(self.tmpdir.name, "existing.csv")
        with open(out, "w") as f:
            f.write("SENTINEL\n")
        with self.assertRaises(SystemExit) as cm:
            main(["extract", src, "-o", out])
        self.assertEqual(cm.exception.code, 1)
        with open(out) as f:
            self.assertIn("SENTINEL", f.read())

    def test_extract_force_overwrites(self):
        from DefFileGenerator.main import main
        src = self._csv_path()
        out = os.path.join(self.tmpdir.name, "existing.csv")
        with open(out, "w") as f:
            f.write("OLD\n")
        main(["extract", src, "-o", out, "--force"])
        with open(out) as f:
            content = f.read()
        self.assertNotEqual(content.strip(), "OLD")

    def test_extract_with_address_offset(self):
        from DefFileGenerator.main import main
        src = self._csv_path("Address,Name,Type\n100,Var1,U16\n")
        out = os.path.join(self.tmpdir.name, "out_offset.csv")
        main(["extract", src, "-o", out, "--address-offset", "10"])
        with open(out) as f:
            content = f.read()
        self.assertIn("110", content)

    def test_extract_stdout(self):
        """extract with no -o writes to stdout."""
        from DefFileGenerator.main import main
        src = self._csv_path()
        captured = io.StringIO()
        with patch("sys.stdout", captured):
            main(["extract", src])
        self.assertIn("Name", captured.getvalue())

    def test_extract_empty_csv_no_data_exits_1(self):
        """Extracting a file with no register data should exit(1)."""
        from DefFileGenerator.main import main
        src = self._csv_path("Address,Name,Type\n")  # header only, no rows
        with self.assertRaises(SystemExit) as cm:
            main(["extract", src])
        self.assertEqual(cm.exception.code, 1)

    # ------------------------------------------------------------------
    # generate
    # ------------------------------------------------------------------
    def test_generate_missing_manufacturer_exits_1(self):
        from DefFileGenerator.main import main
        src = self._csv_path()
        with self.assertRaises(SystemExit) as cm:
            main(["generate", src, "--model", "X"])
        self.assertEqual(cm.exception.code, 1)

    def test_generate_missing_model_exits_1(self):
        from DefFileGenerator.main import main
        src = self._csv_path()
        with self.assertRaises(SystemExit) as cm:
            main(["generate", src, "--manufacturer", "Acme"])
        self.assertEqual(cm.exception.code, 1)

    def test_generate_missing_input_file_exits_1(self):
        from DefFileGenerator.main import main
        with self.assertRaises(SystemExit) as cm:
            main(["generate", "--manufacturer", "X", "--model", "Y"])
        self.assertEqual(cm.exception.code, 1)

    def test_generate_nonexistent_input_exits_1(self):
        from DefFileGenerator.main import main
        with self.assertRaises(SystemExit) as cm:
            main(["generate", "/no/such/file.csv", "--manufacturer", "X", "--model", "Y"])
        self.assertEqual(cm.exception.code, 1)

    def test_generate_template_input_mode(self):
        from DefFileGenerator.main import main
        out = os.path.join(self.tmpdir.name, "tmpl.csv")
        main(["generate", "--template", "-o", out])
        self.assertTrue(os.path.exists(out))
        with open(out) as f:
            self.assertIn("Name", f.read())

    def test_generate_template_definition_mode(self):
        from DefFileGenerator.main import main
        out = os.path.join(self.tmpdir.name, "tmpl_def.csv")
        main(["generate", "--template", "--template-mode", "definition", "-o", out])
        self.assertTrue(os.path.exists(out))
        with open(out) as f:
            self.assertIn("modbusRTU", f.read())

    def test_generate_with_protocol_and_category(self):
        from DefFileGenerator.main import main
        src = self._csv_path()
        out = os.path.join(self.tmpdir.name, "out.csv")
        main(["generate", src, "--manufacturer", "M", "--model", "X",
              "--protocol", "modbustcp", "--category", "Battery", "-o", out])
        with open(out, encoding="utf-8-sig") as f:
            content = f.read()
        self.assertIn("modbustcp", content)
        self.assertIn("Battery", content)

    def test_generate_guard_output_prevents_overwrite(self):
        from DefFileGenerator.main import main
        src = self._csv_path()
        out = os.path.join(self.tmpdir.name, "existing.csv")
        with open(out, "w") as f:
            f.write("SENTINEL\n")
        with self.assertRaises(SystemExit) as cm:
            main(["generate", src, "--manufacturer", "M", "--model", "X", "-o", out])
        self.assertEqual(cm.exception.code, 1)

    # ------------------------------------------------------------------
    # run
    # ------------------------------------------------------------------
    def test_run_missing_manufacturer_exits_1(self):
        from DefFileGenerator.main import main
        src = self._csv_path()
        with self.assertRaises(SystemExit) as cm:
            main(["run", src, "--model", "Y"])
        self.assertEqual(cm.exception.code, 1)

    def test_run_template_mode(self):
        from DefFileGenerator.main import main
        out = os.path.join(self.tmpdir.name, "tmpl.csv")
        main(["run", "--template", "-o", out])
        self.assertTrue(os.path.exists(out))

    def test_run_no_validate_flag(self):
        from DefFileGenerator.main import main
        src = self._csv_path()
        out = os.path.join(self.tmpdir.name, "out.csv")
        main(["run", src, "--manufacturer", "M", "--model", "X", "-o", out, "--no-validate"])
        self.assertTrue(os.path.exists(out))

    def test_run_sheet_on_non_excel_warns(self):
        """--sheet on a CSV file should warn but still succeed."""
        import subprocess
        env = dict(os.environ, PYTHONPATH=REPO_ROOT)
        src = self._csv_path()
        out = os.path.join(self.tmpdir.name, "out.csv")
        result = subprocess.run(
            [sys.executable, "-m", "DefFileGenerator.main",
             "run", src, "--manufacturer", "M", "--model", "X",
             "-o", out, "--sheet", "Sheet1"],
            capture_output=True, text=True, env=env, cwd=REPO_ROOT,
        )
        # Should succeed
        self.assertEqual(result.returncode, 0)
        # Warning must be in stderr
        self.assertIn("sheet", result.stderr.lower())

    # ------------------------------------------------------------------
    # setup_logging
    # ------------------------------------------------------------------
    def test_setup_logging_verbose(self):
        from DefFileGenerator.main import setup_logging
        setup_logging(verbose=True)
        self.assertEqual(logging.getLogger().level, logging.DEBUG)

    def test_setup_logging_quiet(self):
        from DefFileGenerator.main import setup_logging
        setup_logging(quiet=True)
        self.assertEqual(logging.getLogger().level, logging.WARNING)

    def test_setup_logging_default(self):
        from DefFileGenerator.main import setup_logging
        setup_logging()
        self.assertEqual(logging.getLogger().level, logging.INFO)

    # ------------------------------------------------------------------
    # KeyboardInterrupt -> exit 130
    # ------------------------------------------------------------------
    def test_keyboard_interrupt_exits_130(self):
        from DefFileGenerator.main import main
        with patch("DefFileGenerator.main._run_cli", side_effect=KeyboardInterrupt):
            with self.assertRaises(SystemExit) as cm:
                main([])
        self.assertEqual(cm.exception.code, 130)

    # ------------------------------------------------------------------
    # Unexpected exception -> exit 1
    # ------------------------------------------------------------------
    def test_unexpected_exception_exits_1(self):
        from DefFileGenerator.main import main
        with patch("DefFileGenerator.main._run_cli", side_effect=RuntimeError("boom")):
            with self.assertRaises(SystemExit) as cm:
                main([])
        self.assertEqual(cm.exception.code, 1)


if __name__ == "__main__":
    unittest.main()