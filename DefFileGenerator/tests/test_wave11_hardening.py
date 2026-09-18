#!/usr/bin/env python3
"""
Wave 11 Hardening and Edge Case Test Suite for DefFileGenerator.

Provides targeted tests for:
- STR<N> register count calculations in Generator.get_register_count.
- Non-string address handling in Generator.validate_address.
- High-throughput address overlap checking with STR20 and STR40 registers.
- Non-string / integer address validation.
"""

import unittest

from DefFileGenerator.def_gen import Generator


class TestWave11Hardening(unittest.TestCase):
    """Hardening test cases for wave 11 domain fixes."""

    def test_get_register_count_str_types(self) -> None:
        """Verify register count calculations for STR<N> and STRING types."""
        # STR20 is 20 bytes -> 10 16-bit Modbus registers
        self.assertEqual(Generator.get_register_count("STR20", "40001"), 10)
        # STR32 is 32 bytes -> 16 registers
        self.assertEqual(Generator.get_register_count("STR32", "40001"), 16)
        # STRING with _20 suffix -> 10 registers
        self.assertEqual(Generator.get_register_count("STRING", "40001_20"), 10)
        # BITS -> 1 register
        self.assertEqual(Generator.get_register_count("BITS", "40001_0_16"), 1)
        # U32 -> 2 registers
        self.assertEqual(Generator.get_register_count("U32", "40001"), 2)

    def test_validate_address_non_string_input(self) -> None:
        """Verify Generator.validate_address handles integers and None safely."""
        self.assertTrue(Generator.validate_address(40001, "U16"))
        self.assertTrue(Generator.validate_address(0, "U16"))
        self.assertFalse(Generator.validate_address(None, "U16"))
        self.assertFalse(Generator.validate_address(70000, "U16", strict=True))

    def test_str_address_overlap_detection(self) -> None:
        """Verify address overlap detection correctly accounts for STR20 register spans."""
        gen = Generator()
        usage: dict = {}
        warned: set = set()

        # Line 2: STR20 at 40001 occupies 40001..40010 (10 registers)
        overlap1 = gen._check_address_overlap(
            "3", "40001_20", "STRING", "SerialNum", 2, usage, warned
        )
        self.assertFalse(overlap1)

        # Line 3: U16 at 40005 collides with 40001..40010
        overlap2 = gen._check_address_overlap("3", "40005", "U16", "Voltage", 3, usage, warned)
        self.assertTrue(overlap2)

        # Line 4: U16 at 40011 is after 40010, so no collision
        overlap3 = gen._check_address_overlap("3", "40011", "U16", "Current", 4, usage, warned)
        self.assertFalse(overlap3)


if __name__ == "__main__":
    unittest.main()
