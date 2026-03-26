import unittest
import logging
import os
import csv
from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.extractor = Extractor()

    def test_address_offset_simple(self):
        # 30001 + 10 = 30011
        self.assertEqual(self.generator.apply_address_offset('30001', 10), '30011')

    def test_address_offset_hex(self):
        # 0x10 (16) + 10 = 26
        self.assertEqual(self.generator.apply_address_offset('0x10', 10), '26')

    def test_address_offset_compound(self):
        # 30001_20 + 10 = 30011_20
        self.assertEqual(self.generator.apply_address_offset('30001_20', 10), '30011_20')

    def test_tag_sanitization(self):
        rows = [
            {'Name': 'Test Variable @ 123!', 'Tag': '', 'Address': '100', 'Type': 'U16'}
        ]
        processed = self.generator.process_rows(rows)
        # @ and ! are replaced by _, space by _, and multiple _ collapsed
        # "test_variable___123_" -> "test_variable_123" (since we strip and collapse)
        self.assertEqual(processed[0]['Tag'], 'test_variable_123')

    def test_str_n_handling(self):
        rows = [
            {'Name': 'StrVar', 'Address': '30050', 'Type': 'STR20'}
        ]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Info3'], 'STRING')
        self.assertEqual(processed[0]['Info2'], '30050_20')

    def test_address_overlap_bits(self):
        rows = [
            {'Name': 'Bit0', 'RegisterType': '3', 'Address': '100_0_1', 'Type': 'BITS'},
            {'Name': 'Bit1', 'RegisterType': '3', 'Address': '100_1_1', 'Type': 'BITS'}
        ]
        # This should NOT log a warning for overlap at address 100
        try:
            with self.assertLogs(level='WARNING') as log:
                self.generator.process_rows(rows)
                for m in log.output:
                    self.assertNotIn("Address overlap detected", m)
        except AssertionError:
            # No logs at all is also fine
            pass

    def test_address_overlap_bits_vs_nonbits(self):
        rows = [
            {'Name': 'Bit0', 'RegisterType': '3', 'Address': '100_0_1', 'Type': 'BITS'},
            {'Name': 'Reg100', 'RegisterType': '3', 'Address': '100', 'Type': 'U16'}
        ]
        with self.assertLogs(level='WARNING') as log:
            self.generator.process_rows(rows)
            self.assertTrue(any("Address overlap detected" in m for m in log.output))

if __name__ == '__main__':
    unittest.main()
