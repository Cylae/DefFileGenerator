"""Batch 5: Address normalisation and validation edge cases."""

import logging
import unittest

from DefFileGenerator.def_gen import Generator


class TestNormalizeAddressValExtended(unittest.TestCase):
    """Exhaustive normalize_address_val scenarios."""

    def _norm(self, val):
        return Generator.normalize_address_val(val)

    # decimal
    def test_plain_zero(self):
        self.assertEqual(self._norm("0"), "0")

    def test_plain_positive(self):
        self.assertEqual(self._norm("12345"), "12345")

    def test_negative_decimal(self):
        self.assertEqual(self._norm("-1"), "-1")

    # 0x-prefixed hex
    def test_0x_lowercase(self):
        self.assertEqual(self._norm("0xff"), "255")

    def test_0x_uppercase(self):
        self.assertEqual(self._norm("0xFF"), "255")

    def test_0x_zero(self):
        self.assertEqual(self._norm("0x0"), "0")

    def test_0x_large(self):
        self.assertEqual(self._norm("0xFFFF"), "65535")

    def test_0x_invalid_suffix_returned_unchanged(self):
        # "0xGHI" is invalid hex; returned as-is
        result = self._norm("0xGHI")
        self.assertEqual(result, "0xGHI")

    # h-suffixed hex
    def test_h_suffix_lowercase(self):
        self.assertEqual(self._norm("10h"), "16")

    def test_h_suffix_uppercase(self):
        self.assertEqual(self._norm("FFh"), "255")

    def test_h_suffix_invalid(self):
        result = self._norm("GGh")
        self.assertEqual(result, "GGh")

    # bare hex word
    def test_bare_hex_A0(self):
        self.assertEqual(self._norm("A0"), "160")

    def test_bare_hex_FFFF(self):
        self.assertEqual(self._norm("FFFF"), "65535")

    # thousands separator
    def test_thousands_comma(self):
        self.assertEqual(self._norm("1,234"), "1234")

    def test_thousands_no_comma(self):
        self.assertEqual(self._norm("1234"), "1234")

    # leading zeros -> octal / hex interpretation
    def test_leading_zero_bare_hex(self):
        # "012" -> not plain decimal; int("012", 0) raises ValueError in Python3
        # (only "0o12" is octal); falls through to bare hex: int("012", 16) = 18
        self.assertEqual(self._norm("012"), "18")

    def test_leading_zeros_hex(self):
        # "0021" -> not plain decimal, bare hex -> 33
        self.assertEqual(self._norm("0021"), "33")

    # whitespace
    def test_strips_whitespace(self):
        self.assertEqual(self._norm("  100  "), "100")

    # empty / None
    def test_empty_string(self):
        self.assertEqual(self._norm(""), "")

    def test_whitespace_only(self):
        # strip -> empty -> returns ""
        self.assertEqual(self._norm("   "), "")


class TestValidateAddressExtended(unittest.TestCase):
    """validate_address for all compound and special types."""

    def setUp(self):
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)

    # STRING
    def test_string_compound_valid(self):
        self.assertTrue(Generator.validate_address("100_20", "STRING"))

    def test_string_no_underscore_invalid(self):
        self.assertFalse(Generator.validate_address("100", "STRING"))

    def test_str_n_treated_as_string(self):
        self.assertTrue(Generator.validate_address("100_20", "STR20"))

    # BITS
    def test_bits_compound_valid(self):
        self.assertTrue(Generator.validate_address("100_3_1", "BITS"))

    def test_bits_simple_invalid(self):
        self.assertFalse(Generator.validate_address("100", "BITS"))

    # integer types
    def test_u16_plain_decimal(self):
        self.assertTrue(Generator.validate_address("100", "U16"))

    def test_u32_plain_decimal(self):
        self.assertTrue(Generator.validate_address("200", "U32"))

    def test_u16_hex_0x(self):
        self.assertTrue(Generator.validate_address("0x10", "U16"))

    def test_u16_hex_h(self):
        self.assertTrue(Generator.validate_address("10h", "U16"))

    def test_u16_max_valid(self):
        self.assertTrue(Generator.validate_address("65535", "U16"))

    def test_u16_over_max_strict_invalid(self):
        self.assertFalse(Generator.validate_address("65536", "U16", strict=True))

    def test_u16_over_max_lenient_valid(self):
        # Strict=False means out-of-range only warns, not fails
        self.assertTrue(Generator.validate_address("65536", "U16", strict=False))

    def test_negative_address_invalid(self):
        self.assertFalse(Generator.validate_address("-1", "U16"))

    def test_pure_text_invalid(self):
        self.assertFalse(Generator.validate_address("xyz", "U16"))

    # IP / MAC / IPV6 – treated as integer type addresses
    def test_ip_type_decimal_address(self):
        self.assertTrue(Generator.validate_address("100", "IP"))

    def test_mac_type_decimal_address(self):
        self.assertTrue(Generator.validate_address("100", "MAC"))

    def test_ipv6_type_decimal_address(self):
        self.assertTrue(Generator.validate_address("100", "IPV6"))


class TestGetRegisterCountEdge(unittest.TestCase):

    def test_f32_count(self):
        self.assertEqual(Generator.get_register_count("F32", "100"), 2)

    def test_f64_count(self):
        self.assertEqual(Generator.get_register_count("F64", "100"), 4)

    def test_i8_count(self):
        self.assertEqual(Generator.get_register_count("I8", "100"), 1)

    def test_u8_count(self):
        self.assertEqual(Generator.get_register_count("U8", "100"), 1)

    def test_bits_count(self):
        self.assertEqual(Generator.get_register_count("BITS", "100_0_1"), 1)

    def test_ip_count(self):
        self.assertEqual(Generator.get_register_count("IP", "100"), 2)

    def test_string_no_underscore_returns_zero(self):
        self.assertEqual(Generator.get_register_count("STRING", "100"), 0)

    def test_string_odd_length_ceil(self):
        # 11 bytes -> ceil(11/2) = 6
        self.assertEqual(Generator.get_register_count("STRING", "100_11"), 6)

    def test_string_even_length(self):
        # 20 bytes -> ceil(20/2) = 10
        self.assertEqual(Generator.get_register_count("STRING", "100_20"), 10)

    def test_unknown_type_defaults_to_one(self):
        self.assertEqual(Generator.get_register_count("UNKNOWN", "100"), 1)

    def test_u16_wb_count(self):
        self.assertEqual(Generator.get_register_count("U16_WB", "100"), 1)

    def test_u32_wb_count(self):
        self.assertEqual(Generator.get_register_count("U32_WB", "100"), 2)

    def test_f32_wb_count(self):
        self.assertEqual(Generator.get_register_count("F32_WB", "100"), 2)

    def test_u64_wb_count(self):
        self.assertEqual(Generator.get_register_count("U64_WB", "100"), 4)


if __name__ == "__main__":
    unittest.main()