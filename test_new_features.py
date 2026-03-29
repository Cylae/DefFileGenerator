#!/usr/bin/env python3
import unittest
import csv
import os
import io
import logging
from DefFileGenerator.def_gen import Generator

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        # Suppress logging for cleaner test output
        logging.getLogger().setLevel(logging.ERROR)

    def test_address_offset_simple(self):
        rows = [{'Name': 'Var1', 'Address': '100', 'Type': 'U16'}]
        # Offset +10 should result in 110
        processed = self.generator.process_rows(rows, address_offset=10)
        self.assertEqual(processed[0]['Info2'], '110')

    def test_address_offset_compound(self):
        # STRING at 100 with length 10
        rows = [{'Name': 'Str1', 'Address': '100_10', 'Type': 'STRING'}]
        # Offset +50 should result in 150_10
        processed = self.generator.process_rows(rows, address_offset=50)
        self.assertEqual(processed[0]['Info2'], '150_10')

    def test_address_offset_hex(self):
        # Hex address 0x10 (16)
        rows = [{'Name': 'HexVar', 'Address': '0x10', 'Type': 'U16'}]
        # Offset +10 should result in 26
        processed = self.generator.process_rows(rows, address_offset=10)
        self.assertEqual(processed[0]['Info2'], '26')

    def test_str_n_expansion(self):
        # STR20 should become STRING type with length-appended address
        rows = [{'Name': 'StrVar', 'Address': '200', 'Type': 'STR20'}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Info3'], 'STRING')
        self.assertEqual(processed[0]['Info2'], '200_20')

    def test_tag_sanitization(self):
        # Name with spaces and special characters
        rows = [{'Name': 'My Var #1 @ Test!', 'Address': '300', 'Type': 'U16'}]
        processed = self.generator.process_rows(rows)
        # Should be my_var_1_test
        self.assertEqual(processed[0]['Tag'], 'my_var_1_test')

    def test_address_overlap_o1_lookup(self):
        # This is more about logic than speed, but we can verify it detects overlap correctly
        rows = [
            {'Name': 'Var1', 'Address': '400', 'Type': 'U32'}, # Uses 400, 401
            {'Name': 'Var2', 'Address': '401', 'Type': 'U16'}  # Overlaps at 401
        ]
        with self.assertLogs(level='WARNING') as cm:
            self.generator.process_rows(rows)
            # Find overlap message
            self.assertTrue(any("Address overlap detected" in output for output in cm.output))

    def test_bits_overlap_allowed(self):
        # BITS on same address should NOT trigger warning if they are both BITS and have same start
        rows = [
            {'Name': 'Bit1', 'Address': '500_0_1', 'Type': 'BITS'},
            {'Name': 'Bit2', 'Address': '500_1_1', 'Type': 'BITS'}
        ]
        # Using a custom log handler to check for NO warnings
        logger = logging.getLogger()
        handler = io.StringIO()
        ch = logging.StreamHandler(handler)
        logger.addHandler(ch)
        self.generator.process_rows(rows)
        logger.removeHandler(ch)
        self.assertNotIn("Address overlap detected", handler.getvalue())

    def test_negative_address_warning(self):
        rows = [{'Name': 'NegVar', 'Address': '10', 'Type': 'U16'}]
        # Offset -20 should result in -10 and a warning
        with self.assertLogs(level='WARNING') as cm:
            processed = self.generator.process_rows(rows, address_offset=-20)
            self.assertEqual(processed[0]['Info2'], '-10')
            self.assertTrue(any("results in negative address -10" in output for output in cm.output))

if __name__ == "__main__":
    unittest.main()
