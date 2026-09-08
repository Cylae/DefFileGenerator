"""Batch 3: def_gen.py edge-case tests."""

import logging
import math
import tempfile
import os
import unittest

from DefFileGenerator.def_gen import (
    Generator,
    GeneratorConfig,
    generate_template,
    peek_generator,
    run_generator,
)


class TestParseNumeric(unittest.TestCase):
    """_parse_numeric handles all special numeric forms."""

    def _parse(self, val, default=0.0):
        return Generator._parse_numeric(val, default)

    def test_fraction(self):
        self.assertAlmostEqual(self._parse("1/10"), 0.1)

    def test_fraction_zero_denominator(self):
        self.assertEqual(self._parse("1/0"), 0.0)

    def test_comma_as_decimal_separator(self):
        self.assertAlmostEqual(self._parse("1,5"), 1.5)

    def test_thousands_separator_comma(self):
        self.assertAlmostEqual(self._parse("1,000"), 1000.0)

    def test_comma_and_dot_european(self):
        # European: 1.234,56 -> 1234.56
        self.assertAlmostEqual(self._parse("1.234,56"), 1234.56)

    def test_comma_and_dot_us(self):
        # US: 1,234.56 -> 1234.56
        self.assertAlmostEqual(self._parse("1,234.56"), 1234.56)

    def test_empty_string_returns_default(self):
        self.assertEqual(self._parse("", default=42.0), 42.0)

    def test_none_returns_default(self):
        self.assertEqual(self._parse(None, default=7.0), 7.0)

    def test_whitespace_only_returns_default(self):
        self.assertEqual(self._parse("   ", default=3.0), 3.0)

    def test_non_numeric_returns_default(self):
        self.assertEqual(self._parse("abc", default=-1.0), -1.0)

    def test_integer(self):
        self.assertEqual(self._parse("42"), 42.0)

    def test_negative_float(self):
        self.assertAlmostEqual(self._parse("-3.14"), -3.14)


class TestCalculateCoefficients(unittest.TestCase):
    """_calculate_coefficients applies scale factors correctly."""

    def test_no_scale(self):
        a, b = Generator._calculate_coefficients("2", "5", "")
        self.assertEqual(a, "2.000000")
        self.assertEqual(b, "5.000000")

    def test_scale_1(self):
        a, b = Generator._calculate_coefficients("1", "0", "1")
        self.assertAlmostEqual(float(a), 10.0)

    def test_scale_2(self):
        a, b = Generator._calculate_coefficients("0.5", "0", "2")
        self.assertAlmostEqual(float(a), 50.0)

    def test_invalid_scale_treated_as_zero(self):
        a, b = Generator._calculate_coefficients("1", "0", "invalid")
        self.assertAlmostEqual(float(a), 1.0)

    def test_none_inputs(self):
        a, b = Generator._calculate_coefficients(None, None, None)
        self.assertAlmostEqual(float(a), 1.0)
        self.assertAlmostEqual(float(b), 0.0)


class TestApplyAddressOffset(unittest.TestCase):
    """apply_address_offset handles edge cases."""

    def test_simple_offset(self):
        self.assertEqual(Generator.apply_address_offset("100", 10), "110")

    def test_zero_offset(self):
        self.assertEqual(Generator.apply_address_offset("100", 0), "100")

    def test_compound_string_address(self):
        self.assertEqual(Generator.apply_address_offset("100_20", 5), "105_20")

    def test_compound_bits_address(self):
        self.assertEqual(Generator.apply_address_offset("100_3_1", 5), "105_3_1")

    def test_hex_base(self):
        result = Generator.apply_address_offset("0x10", 10)
        self.assertEqual(result, "26")

    def test_empty_address(self):
        self.assertEqual(Generator.apply_address_offset("", 10), "")

    def test_negative_result_is_valid_str(self):
        result = Generator.apply_address_offset("5", -10)
        self.assertIn("-5", result)

    def test_none_address(self):
        self.assertEqual(Generator.apply_address_offset(None, 10), "")


class TestNormalizeTypeExtended(unittest.TestCase):
    """normalize_type edge cases beyond the precedence test."""

    def test_ip_type(self):
        self.assertEqual(Generator.normalize_type("IP"), "IP")

    def test_ipv6_type(self):
        self.assertEqual(Generator.normalize_type("IPV6"), "IPV6")

    def test_mac_type(self):
        self.assertEqual(Generator.normalize_type("MAC"), "MAC")

    def test_bits_type(self):
        self.assertEqual(Generator.normalize_type("BITS"), "BITS")

    def test_type_with_spaces_stripped(self):
        self.assertEqual(Generator.normalize_type("  u16  "), "U16")

    def test_none_type(self):
        self.assertEqual(Generator.normalize_type(None), "U16")

    def test_empty_type(self):
        self.assertEqual(Generator.normalize_type(""), "U16")

    def test_uint32_wb_suffix(self):
        self.assertEqual(Generator.normalize_type("uint32_wb"), "U32_WB")

    def test_int16_b_suffix(self):
        self.assertEqual(Generator.normalize_type("int16 big"), "I16_B")

    def test_float32_word_suffix(self):
        self.assertEqual(Generator.normalize_type("float32 word"), "F32_W")


class TestValidateTypeExtended(unittest.TestCase):

    def test_ip_valid(self):
        self.assertTrue(Generator.validate_type("IP"))

    def test_ipv6_valid(self):
        self.assertTrue(Generator.validate_type("IPV6"))

    def test_mac_valid(self):
        self.assertTrue(Generator.validate_type("MAC"))

    def test_bits_valid(self):
        self.assertTrue(Generator.validate_type("BITS"))

    def test_string_valid(self):
        self.assertTrue(Generator.validate_type("STRING"))

    def test_str_n_valid(self):
        self.assertTrue(Generator.validate_type("STR10"))
        self.assertTrue(Generator.validate_type("STR100"))

    def test_u16_with_suffix_valid(self):
        self.assertTrue(Generator.validate_type("U16_WB"))
        self.assertTrue(Generator.validate_type("U16_B"))
        self.assertTrue(Generator.validate_type("U16_W"))

    def test_f32_valid(self):
        self.assertTrue(Generator.validate_type("F32"))

    def test_f64_valid(self):
        self.assertTrue(Generator.validate_type("F64"))

    def test_invalid_type(self):
        self.assertFalse(Generator.validate_type("GARBAGE"))

    def test_integer_string_invalid(self):
        self.assertFalse(Generator.validate_type("123"))


class TestProcessRowsEdgeCases(unittest.TestCase):

    def setUp(self):
        self.g = Generator()
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)

    def test_skips_row_with_both_name_and_address_empty(self):
        rows = [{"Name": "", "Address": "", "Type": "U16"}]
        result = list(self.g.process_rows(rows))
        self.assertEqual(result, [])

    def test_skips_row_with_invalid_type(self):
        rows = [{"Name": "X", "Address": "1", "Type": "GARBAGE"}]
        result = list(self.g.process_rows(rows))
        self.assertEqual(result, [])

    def test_skips_row_with_invalid_address(self):
        rows = [{"Name": "X", "Address": "70000", "Type": "U16"}]
        result = list(self.g.process_rows(rows))
        self.assertEqual(result, [])

    def test_str_type_becomes_string_with_compound_address(self):
        rows = [{"Name": "S", "Address": "100", "Type": "STR20"}]
        result = list(self.g.process_rows(rows))
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["Info3"], "STRING")
        self.assertIn("_", result[0]["Info2"])

    def test_str_type_with_existing_underscore_not_doubled(self):
        rows = [{"Name": "S", "Address": "100_20", "Type": "STR20"}]
        result = list(self.g.process_rows(rows))
        self.assertEqual(result[0]["Info2"], "100_20")

    def test_all_empty_row_skipped_silently(self):
        rows = [{"Name": "", "Address": "", "Type": "", "Unit": ""}]
        result = list(self.g.process_rows(rows))
        self.assertEqual(result, [])

    def test_address_offset_applied_in_process_rows(self):
        rows = [{"Name": "V", "Address": "100", "Type": "U16"}]
        result = list(self.g.process_rows(rows, address_offset=50))
        self.assertEqual(result[0]["Info2"], "150")

    def test_coef_a_and_b_defaults_when_empty(self):
        rows = [{"Name": "V", "Address": "100", "Type": "U16",
                 "Factor": "", "Offset": "", "ScaleFactor": ""}]
        result = list(self.g.process_rows(rows))
        self.assertEqual(result[0]["CoefA"], "1.000000")
        self.assertEqual(result[0]["CoefB"], "0.000000")

    def test_unit_preserved(self):
        rows = [{"Name": "V", "Address": "100", "Type": "U16", "Unit": "kWh"}]
        result = list(self.g.process_rows(rows))
        self.assertEqual(result[0]["Unit"], "kWh")

    def test_tag_auto_generated_numeric_prefix_fixed(self):
        rows = [{"Name": "123 Var", "Address": "100", "Type": "U16"}]
        result = list(self.g.process_rows(rows))
        tag = result[0]["Tag"]
        self.assertTrue(tag[0].isalpha() or tag.startswith("v_"), f"Bad tag: {tag}")


class TestDetermineInfo1Extended(unittest.TestCase):

    def setUp(self):
        self.g = Generator()

    def test_numeric_code_3(self):
        self.assertEqual(self.g._determine_info1("3"), "3")

    def test_numeric_code_1(self):
        self.assertEqual(self.g._determine_info1("1"), "1")

    def test_empty_string_returns_3(self):
        self.assertEqual(self.g._determine_info1(""), "3")

    def test_none_returns_3(self):
        self.assertEqual(self.g._determine_info1(None), "3")

    def test_unknown_string_returns_3(self):
        logging.disable(logging.CRITICAL)
        self.assertEqual(self.g._determine_info1("bogus"), "3")
        logging.disable(logging.NOTSET)


class TestGenerateTemplate(unittest.TestCase):

    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.tmpdir.cleanup()

    def test_input_mode_creates_file(self):
        out = os.path.join(self.tmpdir.name, "tmpl_in.csv")
        generate_template(out, mode="input")
        with open(out) as f:
            content = f.read()
        self.assertIn("Name", content)
        self.assertIn("RegisterType", content)

    def test_definition_mode_creates_file(self):
        out = os.path.join(self.tmpdir.name, "tmpl_def.csv")
        generate_template(out, mode="definition")
        with open(out, encoding="utf-8-sig") as f:
            content = f.read()
        self.assertIn("modbusRTU", content)

    def test_stdout_mode_no_crash(self):
        import io
        from unittest.mock import patch
        buf = io.StringIO()
        with patch("sys.stdout", buf):
            generate_template(None, mode="input")
        self.assertIn("Name", buf.getvalue())


class TestRunGenerator(unittest.TestCase):

    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        self.tmpdir.cleanup()
        logging.disable(logging.NOTSET)

    def _csv(self, content="Name,Address,Type\nVar1,100,U16\n"):
        p = os.path.join(self.tmpdir.name, "in.csv")
        with open(p, "w", encoding="utf-8") as f:
            f.write(content)
        return p

    def test_run_generator_with_input_file(self):
        src = self._csv()
        out = os.path.join(self.tmpdir.name, "out.csv")
        config = GeneratorConfig(
            input_file=src, output=out,
            manufacturer="ACME", model="X1",
        )
        run_generator(config)
        self.assertTrue(os.path.exists(out))

    def test_run_generator_with_input_data(self):
        out = os.path.join(self.tmpdir.name, "out2.csv")
        config = GeneratorConfig(output=out, manufacturer="M", model="Y")
        rows = [{"Name": "V", "Address": "10", "Type": "U16"}]
        run_generator(config, input_data=rows)
        self.assertTrue(os.path.exists(out))

    def test_run_generator_template(self):
        out = os.path.join(self.tmpdir.name, "tmpl.csv")
        config = GeneratorConfig(template=True, output=out)
        run_generator(config)
        self.assertTrue(os.path.exists(out))

    def test_run_generator_no_input_file_logs_error(self):
        config = GeneratorConfig(manufacturer="M", model="X")
        # Should not raise; just log error
        run_generator(config)

    def test_run_generator_nonexistent_file_logs_error(self):
        config = GeneratorConfig(
            input_file="/no/such.csv", manufacturer="M", model="X"
        )
        run_generator(config)

    def test_run_generator_semicolon_delimited_input(self):
        p = os.path.join(self.tmpdir.name, "semi.csv")
        with open(p, "w", encoding="utf-8") as f:
            f.write("Name;Address;Type\nVar1;100;U16\n")
        out = os.path.join(self.tmpdir.name, "out_semi.csv")
        config = GeneratorConfig(input_file=p, output=out, manufacturer="M", model="X")
        run_generator(config)
        self.assertTrue(os.path.exists(out))


class TestPeekGenerator(unittest.TestCase):

    def test_non_empty(self):
        has, it = peek_generator([1, 2, 3])
        self.assertTrue(has)
        self.assertEqual(list(it), [1, 2, 3])

    def test_empty_iterable(self):
        has, it = peek_generator([])
        self.assertFalse(has)
        self.assertEqual(list(it), [])

    def test_none_input(self):
        has, it = peek_generator(None)
        self.assertFalse(has)
        self.assertEqual(list(it), [])

    def test_generator_input(self):
        has, it = peek_generator(x for x in range(3))
        self.assertTrue(has)
        self.assertEqual(list(it), [0, 1, 2])


if __name__ == "__main__":
    unittest.main()