"""Wave-2 targeted coverage tests.

Covers the specific branches identified as missing from the previous coverage run:

def_gen.py gaps
---------------
* _parse_numeric: fraction strings, locale decimals (comma-dot, dot-comma), thousands
* apply_address_offset: negative result triggers warning; non-numeric base is passed through
* _calculate_coefficients: scale_factor applied correctly; bad scale_factor defaults to 0
* validate_csv: strict overlap failure; lenient overlap is only a warning; invalid header;
  non-existing path
* write_output_csv: output=None -> stdout; output=file-object; output=file path
* run_generator: template with input_data uses definition mode; no input_file logs error;
  non-existent input_file logs error

extractor.py gaps
-----------------
* extract_from_csv: semicolon-delimited file; empty file (no rows); UTF-16 encoded file
* extract_from_xml: valid XML; file not found
* map_and_clean: empty tables skipped; None tables skipped
* Extractor: custom mapping overrides defaults

main.py gaps
------------
* --pages on non-PDF file emits a WARNING
* --sheet on non-Excel file emits a WARNING
* validate command: valid file succeeds; non-existent file exits 1; bad file exits 1

sanitize_csv_field security
---------------------------
* Formula injection triggers and safe numeric values
"""

import csv
import io
import logging
import os
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

from DefFileGenerator.def_gen import GeneratorConfig, Generator, run_generator
from DefFileGenerator.extractor import Extractor


REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_def_csv(rows, header=None):
    """Return a valid semicolon-delimited WebdynSunPM definition CSV as a string."""
    lines = []
    hdr = header if header is not None else [
        "modbusRTU", "Inverter", "TestMfr", "TestModel", "", "", "", "", "", "", ""
    ]
    lines.append(";".join(str(x) for x in hdr))
    for row in rows:
        lines.append(";".join(str(x) for x in row))
    return "\n".join(lines) + "\n"


# ===========================================================================
# _parse_numeric edge cases
# ===========================================================================

class TestParseNumeric(unittest.TestCase):
    def _pn(self, val, default=0.0):
        return Generator._parse_numeric(val, default)

    def test_none_returns_default(self):
        self.assertAlmostEqual(self._pn(None), 0.0)

    def test_empty_string_returns_default(self):
        self.assertAlmostEqual(self._pn(""), 0.0)

    def test_whitespace_returns_default(self):
        self.assertAlmostEqual(self._pn("   "), 0.0)

    def test_simple_float(self):
        self.assertAlmostEqual(self._pn("3.14"), 3.14)

    def test_fraction_string(self):
        self.assertAlmostEqual(self._pn("1/3"), 1 / 3, places=9)

    def test_fraction_zero_denominator_returns_default(self):
        self.assertAlmostEqual(self._pn("5/0"), 0.0)

    def test_fraction_bad_parts_returns_default(self):
        self.assertAlmostEqual(self._pn("a/b"), 0.0)

    def test_comma_dot_european_thousands(self):
        # "1,234.56" - comma before dot -> strip comma -> 1234.56
        self.assertAlmostEqual(self._pn("1,234.56"), 1234.56)

    def test_dot_comma_european_decimal(self):
        # "1.234,56" - dot before comma -> European -> 1234.56
        self.assertAlmostEqual(self._pn("1.234,56"), 1234.56)

    def test_comma_only_thousands(self):
        # "1,000,000" matches ^-?\d{1,3}(,\d{3})+ pattern -> strip commas
        self.assertAlmostEqual(self._pn("1,000,000"), 1_000_000.0)

    def test_comma_decimal_separator(self):
        # "3,14" - single comma, not thousands pattern -> replace , with .
        self.assertAlmostEqual(self._pn("3,14"), 3.14)

    def test_integer_string(self):
        self.assertAlmostEqual(self._pn("100"), 100.0)

    def test_negative_value(self):
        self.assertAlmostEqual(self._pn("-42.5"), -42.5)

    def test_bad_string_returns_default(self):
        self.assertAlmostEqual(self._pn("N/A"), 0.0)
        self.assertAlmostEqual(self._pn("not-a-number"), 0.0)


# ===========================================================================
# apply_address_offset
# ===========================================================================

class TestApplyAddressOffset(unittest.TestCase):
    def test_empty_address_returns_empty(self):
        result = Generator.apply_address_offset("", 10)
        self.assertEqual(result, "")

    def test_none_address_returns_empty(self):
        result = Generator.apply_address_offset(None, 10)
        self.assertEqual(result, "")

    def test_positive_offset(self):
        self.assertEqual(Generator.apply_address_offset("100", 10), "110")

    def test_negative_offset_valid(self):
        self.assertEqual(Generator.apply_address_offset("100", -10), "90")

    def test_negative_result_logs_warning(self):
        with self.assertLogs(level="WARNING") as cm:
            result = Generator.apply_address_offset("5", -10, name="TestVar")
        self.assertEqual(result, "-5")
        self.assertTrue(any("negative" in m.lower() for m in cm.output))

    def test_compound_address_only_base_shifted(self):
        # "100_2_1" -> bits address: base=100+5=105, rest preserved
        result = Generator.apply_address_offset("100_2_1", 5)
        self.assertEqual(result, "105_2_1")

    def test_non_numeric_base_passed_through(self):
        result = Generator.apply_address_offset("INVALID", 5)
        self.assertIsInstance(result, str)

    def test_zero_offset_is_noop(self):
        self.assertEqual(Generator.apply_address_offset("12345", 0), "12345")

    def test_hex_address_normalized_and_offset(self):
        # "0x10" = 16, + 4 = 20
        result = Generator.apply_address_offset("0x10", 4)
        self.assertEqual(result, "20")


# ===========================================================================
# _calculate_coefficients
# ===========================================================================

class TestCalculateCoefficients(unittest.TestCase):
    def test_simple_factor_offset(self):
        a, b = Generator._calculate_coefficients("2.0", "0.5", "")
        self.assertEqual(a, "2.000000")
        self.assertEqual(b, "0.500000")

    def test_scale_factor_applied(self):
        a, b = Generator._calculate_coefficients("1", "0", "2")
        self.assertEqual(a, "100.000000")

    def test_bad_scale_factor_defaults_to_zero(self):
        a, b = Generator._calculate_coefficients("3", "0", "notanumber")
        self.assertEqual(a, "3.000000")

    def test_none_factor_defaults_one(self):
        a, b = Generator._calculate_coefficients(None, None, "")
        self.assertEqual(a, "1.000000")
        self.assertEqual(b, "0.000000")

    def test_negative_scale(self):
        a, b = Generator._calculate_coefficients("1", "0", "-1")
        self.assertAlmostEqual(float(a), 0.1, places=5)

    def test_fractional_factor(self):
        a, b = Generator._calculate_coefficients("0.001", "0", "")
        self.assertAlmostEqual(float(a), 0.001, places=6)


# ===========================================================================
# validate_csv extended paths
# ===========================================================================

class TestValidateCsvExtended(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.tmpdir = tempfile.TemporaryDirectory()
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        self.tmpdir.cleanup()
        logging.disable(logging.NOTSET)

    def _write(self, name, content, encoding="utf-8"):
        path = os.path.join(self.tmpdir.name, name)
        with open(path, "w", encoding=encoding, newline="") as f:
            f.write(content)
        return path

    def test_nonexistent_file_returns_false(self):
        logging.disable(logging.NOTSET)
        with self.assertLogs(level="ERROR"):
            result = self.generator.validate_csv("/no/such/file.csv")
        self.assertFalse(result)

    def test_invalid_header_returns_false(self):
        path = self._write("bad_hdr.csv", "\n")
        logging.disable(logging.NOTSET)
        with self.assertLogs(level="ERROR"):
            result = self.generator.validate_csv(path)
        self.assertFalse(result)

    def test_strict_overlap_fails(self):
        content = _make_def_csv([
            ["1", "3", "40001", "U32", "", "Var1", "tag1", "1.0", "0.0", "W", "4"],
            ["2", "3", "40002", "U16", "", "Var2", "tag2", "1.0", "0.0", "W", "4"],
        ])
        path = self._write("overlap.csv", content)
        logging.disable(logging.NOTSET)
        result = self.generator.validate_csv(path, strict=True)
        self.assertFalse(result)

    def test_lenient_overlap_passes(self):
        content = _make_def_csv([
            ["1", "3", "40001", "U32", "", "Var1", "tag1", "1.0", "0.0", "W", "4"],
            ["2", "3", "40002", "U16", "", "Var2", "tag2", "1.0", "0.0", "W", "4"],
        ])
        path = self._write("overlap_lenient.csv", content)
        result = self.generator.validate_csv(path, strict=False)
        self.assertTrue(result)

    def test_valid_file_passes(self):
        content = _make_def_csv([
            ["1", "3", "40001", "U16", "", "Var1", "tag1", "1.0", "0.0", "W", "4"],
        ])
        path = self._write("valid.csv", content)
        result = self.generator.validate_csv(path)
        self.assertTrue(result)

    def test_invalid_type_fails(self):
        content = _make_def_csv([
            ["1", "3", "40001", "BADTYPE", "", "Var1", "tag1", "1.0", "0.0", "W", "4"],
        ])
        path = self._write("bad_type.csv", content)
        logging.disable(logging.NOTSET)
        result = self.generator.validate_csv(path)
        self.assertFalse(result)


# ===========================================================================
# write_output_csv: stdout and file-object paths
# ===========================================================================

class TestWriteOutputCsv(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)

    def _rows(self):
        return iter([{
            "Name": "TestVar", "Tag": "test_var",
            "RegisterType": "Holding Register", "Address": "100",
            "Type": "U16", "Factor": "1", "Offset": "0",
            "Unit": "V", "Action": "4", "ScaleFactor": "0",
        }])

    def test_output_none_writes_to_stdout(self):
        rows = list(self.generator.process_rows(self._rows()))
        captured = io.StringIO()
        with patch("sys.stdout", captured):
            Generator.write_output_csv(None, iter(rows), "MFR", "MDL")
        output = captured.getvalue()
        self.assertIn("modbusRTU", output)

    def test_output_file_object(self):
        rows = list(self.generator.process_rows(self._rows()))
        buf = io.StringIO()
        Generator.write_output_csv(buf, iter(rows), "MFR", "MDL")
        output = buf.getvalue()
        self.assertIn("modbusRTU", output)

    def test_output_file_path_creates_file(self):
        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as f:
            path = f.name
        try:
            rows = list(self.generator.process_rows(self._rows()))
            Generator.write_output_csv(path, iter(rows), "MFR", "MDL")
            self.assertTrue(os.path.exists(path))
            with open(path, encoding="utf-8-sig") as f:
                content = f.read()
            self.assertIn("modbusRTU", content)
        finally:
            if os.path.exists(path):
                os.unlink(path)


# ===========================================================================
# run_generator paths
# ===========================================================================

class TestRunGenerator(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        self.tmpdir.cleanup()
        logging.disable(logging.NOTSET)

    def test_template_with_input_data_uses_definition_mode(self):
        out = os.path.join(self.tmpdir.name, "tmpl.csv")
        config = GeneratorConfig(
            input_file=None, output=out,
            manufacturer=None, model=None,
            template=True, template_mode="input",
        )
        input_data = iter([{"dummy": "row"}])
        run_generator(config, input_data=input_data)
        with open(out) as f:
            content = f.read()
        self.assertIn("#Index", content)

    def test_no_input_file_logs_error(self):
        config = GeneratorConfig(
            input_file=None, output=None,
            manufacturer="X", model="Y",
            template=False,
        )
        logging.disable(logging.NOTSET)
        with self.assertLogs(level="ERROR"):
            run_generator(config)

    def test_nonexistent_input_file_logs_error(self):
        config = GeneratorConfig(
            input_file="/no/such/file.csv", output=None,
            manufacturer="X", model="Y",
            template=False,
        )
        logging.disable(logging.NOTSET)
        with self.assertLogs(level="ERROR"):
            run_generator(config)

    def test_run_with_input_data_succeeds(self):
        out = os.path.join(self.tmpdir.name, "out.csv")
        config = GeneratorConfig(
            input_file=None, output=out,
            manufacturer="ACME", model="X100",
            template=False,
        )
        input_data = [{"Name": "Var", "Address": "100", "Type": "U16",
                       "Factor": "1", "Offset": "0", "Unit": "V",
                       "Action": "4", "RegisterType": "Holding Register",
                       "Tag": "", "ScaleFactor": ""}]
        run_generator(config, input_data=iter(input_data))
        self.assertTrue(os.path.exists(out))

    def test_run_from_csv_file_succeeds(self):
        csv_path = os.path.join(self.tmpdir.name, "input.csv")
        out = os.path.join(self.tmpdir.name, "out.csv")
        with open(csv_path, "w", encoding="utf-8", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["Name", "Address", "Type",
                                                    "Factor", "Offset", "Unit",
                                                    "Action", "RegisterType",
                                                    "Tag", "ScaleFactor"])
            writer.writeheader()
            writer.writerow({"Name": "Var1", "Address": "100", "Type": "U16",
                             "Factor": "1", "Offset": "0", "Unit": "V",
                             "Action": "4", "RegisterType": "Holding Register",
                             "Tag": "", "ScaleFactor": ""})
        config = GeneratorConfig(
            input_file=csv_path, output=out,
            manufacturer="ACME", model="X100",
            template=False,
        )
        run_generator(config)
        self.assertTrue(os.path.exists(out))


# ===========================================================================
# Extractor: CSV extraction edge cases
# ===========================================================================

class TestExtractorCsv(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()
        self.tmpdir = tempfile.TemporaryDirectory()
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        self.tmpdir.cleanup()
        logging.disable(logging.NOTSET)

    def _write(self, name, content, encoding="utf-8"):
        path = os.path.join(self.tmpdir.name, name)
        with open(path, "w", encoding=encoding, newline="") as f:
            f.write(content)
        return path

    def test_csv_basic_extraction(self):
        path = self._write("data.csv", "Address,Name,Type\n100,Voltage,U16\n101,Current,I16\n")
        tables = self.extractor.extract_from_csv(path)
        rows = list(self.extractor.map_and_clean(tables))
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["Address"], "100")

    def test_csv_semicolon_delimited(self):
        path = self._write("semi.csv", "Address;Name;Type\n200;Power;U32\n")
        tables = self.extractor.extract_from_csv(path)
        rows = list(self.extractor.map_and_clean(tables))
        self.assertGreaterEqual(len(rows), 1)
        self.assertEqual(rows[0]["Address"], "200")

    def test_csv_empty_header_only(self):
        path = self._write("empty.csv", "Address,Name,Type\n")
        tables = self.extractor.extract_from_csv(path)
        rows = list(self.extractor.map_and_clean(tables))
        self.assertEqual(rows, [])

    def test_csv_utf16_file(self):
        path = self._write("utf16.csv", "Address,Name,Type\n300,Speed,U16\n", encoding="utf-16")
        tables = self.extractor.extract_from_csv(path)
        rows = list(self.extractor.map_and_clean(tables))
        self.assertGreaterEqual(len(rows), 1)
        self.assertEqual(rows[0]["Address"], "300")


# ===========================================================================
# Extractor: XML extraction
# ===========================================================================

class TestExtractorXml(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()
        self.tmpdir = tempfile.TemporaryDirectory()
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        self.tmpdir.cleanup()
        logging.disable(logging.NOTSET)

    def _write(self, name, content):
        path = os.path.join(self.tmpdir.name, name)
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        return path

    def test_valid_xml_extraction(self):
        xml = (
            '<?xml version="1.0"?>\n<registers>\n'
            '  <register><Address>1000</Address><Name>Temperature</Name>'
            '<Type>U16</Type></register>\n'
            '  <register><Address>1001</Address><Name>Humidity</Name>'
            '<Type>U16</Type></register>\n</registers>'
        )
        path = self._write("regs.xml", xml)
        tables = self.extractor.extract_from_xml(path)
        rows = list(self.extractor.map_and_clean(tables))
        self.assertGreaterEqual(len(rows), 1)
        names = [r.get("Name") for r in rows]
        self.assertIn("Temperature", names)

    def test_xml_missing_file_graceful(self):
        tables = self.extractor.extract_from_xml("/no/such/file.xml")
        logging.disable(logging.NOTSET)
        rows = list(self.extractor.map_and_clean(tables))
        self.assertEqual(rows, [])


# ===========================================================================
# Extractor: custom mapping and edge cases
# ===========================================================================

class TestExtractorCustomMapping(unittest.TestCase):
    def test_custom_mapping_overrides_defaults(self):
        custom = {"Name": "Beschrijving", "Address": "Register"}
        extractor = Extractor(custom)
        raw_data = [[{"Beschrijving": "Volt", "Register": "500", "Type": "U16"}]]
        rows = list(extractor.map_and_clean(raw_data))
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Name"], "Volt")
        self.assertEqual(rows[0]["Address"], "500")

    def test_empty_tables_skipped(self):
        extractor = Extractor()
        rows = list(extractor.map_and_clean([iter([])]))
        self.assertEqual(rows, [])

    def test_none_tables_skipped(self):
        extractor = Extractor()
        rows = list(extractor.map_and_clean(None))
        self.assertEqual(rows, [])


# ===========================================================================
# main.py: --pages and --sheet warning on wrong file types
# ===========================================================================

class TestMainWarnings(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.tmpdir.cleanup()

    def _csv_path(self, name="input.csv"):
        path = os.path.join(self.tmpdir.name, name)
        with open(path, "w", encoding="utf-8") as f:
            f.write("Address,Name,Type\n100,Var,U16\n")
        return path

    def test_pages_on_csv_warns(self):
        src = self._csv_path()
        out = os.path.join(self.tmpdir.name, "out.csv")
        env = dict(os.environ, PYTHONPATH=REPO_ROOT)
        result = subprocess.run(
            [sys.executable, "-m", "DefFileGenerator.main",
             "run", src, "--manufacturer", "M", "--model", "X",
             "-o", out, "--pages", "1,2"],
            capture_output=True, text=True, env=env, cwd=REPO_ROOT,
        )
        self.assertIn("pages", result.stderr.lower())

    def test_sheet_on_csv_warns(self):
        src = self._csv_path()
        out = os.path.join(self.tmpdir.name, "out.csv")
        env = dict(os.environ, PYTHONPATH=REPO_ROOT)
        result = subprocess.run(
            [sys.executable, "-m", "DefFileGenerator.main",
             "run", src, "--manufacturer", "M", "--model", "X",
             "-o", out, "--sheet", "Sheet1"],
            capture_output=True, text=True, env=env, cwd=REPO_ROOT,
        )
        self.assertIn("sheet", result.stderr.lower())


# ===========================================================================
# validate command via main()
# ===========================================================================

class TestValidateCommand(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        self.tmpdir.cleanup()
        logging.disable(logging.NOTSET)

    def _valid_csv(self):
        path = os.path.join(self.tmpdir.name, "valid.csv")
        content = _make_def_csv([
            ["1", "3", "40001", "U16", "", "Var1", "tag1", "1.0", "0.0", "V", "4"],
        ])
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        return path

    def test_validate_valid_file_exits_0(self):
        from DefFileGenerator.main import main
        path = self._valid_csv()
        main(["validate", path])  # Should not raise

    def test_validate_nonexistent_file_exits_1(self):
        from DefFileGenerator.main import main
        with self.assertRaises(SystemExit) as cm:
            main(["validate", "/no/such/def.csv"])
        self.assertEqual(cm.exception.code, 1)

    def test_validate_bad_file_exits_1(self):
        from DefFileGenerator.main import main
        path = os.path.join(self.tmpdir.name, "bad.csv")
        content = _make_def_csv([
            ["1", "3", "40001", "U16", "", "Var1", "dup_tag", "1.0", "0.0", "V", "4"],
            ["2", "3", "40003", "U16", "", "Var2", "dup_tag", "1.0", "0.0", "V", "4"],
        ])
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        logging.disable(logging.NOTSET)
        with self.assertRaises(SystemExit) as cm:
            main(["validate", path])
        self.assertEqual(cm.exception.code, 1)


# ===========================================================================
# sanitize_csv_field security
# ===========================================================================

class TestSanitizeCsvField(unittest.TestCase):
    """Verify CSV formula-injection sanitization hardening."""

    def _s(self, val):
        return Generator.sanitize_csv_field(val)

    def test_normal_string_untouched(self):
        self.assertEqual(self._s("hello"), "hello")

    def test_formula_eq_prefixed(self):
        self.assertEqual(self._s("=SUM(A1)"), "'=SUM(A1)")

    def test_at_sign_prefixed(self):
        self.assertEqual(self._s("@EXEC"), "'@EXEC")

    def test_plus_numeric_allowed(self):
        result = self._s("+42.5")
        self.assertFalse(result.startswith("'"), f"Expected no escape, got: {result}")

    def test_minus_numeric_allowed(self):
        result = self._s("-3.14")
        self.assertFalse(result.startswith("'"), f"Expected no escape, got: {result}")

    def test_pipe_prefixed(self):
        self.assertEqual(self._s("|cmd"), "'|cmd")

    def test_tab_prefixed(self):
        self.assertEqual(self._s("\tcmd"), "'\tcmd")

    def test_fullwidth_eq_prefixed(self):
        self.assertEqual(self._s("\uff1d=trick"), "'\uff1d=trick")

    def test_empty_string_untouched(self):
        self.assertEqual(self._s(""), "")

    def test_leading_whitespace_formula_prefixed(self):
        self.assertEqual(self._s("  =cmd"), "'  =cmd")


if __name__ == "__main__":
    unittest.main()
