#!/usr/bin/env python3
import unittest
from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.extractor = Extractor()

    def test_apply_address_offset(self):
        # Simple integer
        self.assertEqual(self.generator.apply_address_offset("30001", 10), "30011")
        # Hex input
        self.assertEqual(self.generator.apply_address_offset("0x10", 5), "21")
        # Compound address (Address_Length)
        self.assertEqual(self.generator.apply_address_offset("40001_20", 100), "40101_20")
        # Compound address (Address_StartBit_Length)
        self.assertEqual(self.generator.apply_address_offset("30001_0_16", 5), "30006_0_16")

    def test_tag_sanitization(self):
        seen_tags = {}
        row = {'Name': 'Test Variable !!!'}
        name, tag = self.generator._process_name_and_tag(row, 2, {}, seen_tags)
        self.assertEqual(tag, "test_variable")

        # Test uniqueness
        row2 = {'Name': 'Test Variable ???'}
        name2, tag2 = self.generator._process_name_and_tag(row2, 3, {}, seen_tags)
        self.assertEqual(tag2, "test_variable_1")

        # Multiple underscores
        row3 = {'Name': 'My   Awesome---Signal'}
        name3, tag3 = self.generator._process_name_and_tag(row3, 4, {}, seen_tags)
        self.assertEqual(tag3, "my_awesome_signal")

    def test_str_n_type(self):
        rows = [
            {'Name': 'TestStr', 'Address': '30050', 'Type': 'STR20', 'RegisterType': 'Holding Register'}
        ]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Info3'], 'STRING')
        self.assertEqual(processed[0]['Info2'], '30050_20')

    def test_address_overlap_o1(self):
        address_usage = {}
        # 40001 (U16)
        self.generator._check_address_overlap('3', '40001', 'U16', 'Var1', 2, address_usage)
        # 40001 (U32) -> Should overlap
        import logging
        with self.assertLogs(level='WARNING') as cm:
            self.generator._check_address_overlap('3', '40001', 'U32', 'Var2', 3, address_usage)
            self.assertTrue(any("Address overlap detected" in output for output in cm.output))

    def test_complex_address_construction(self):
        table = [{
            'Name': 'BitVar',
            'Address': '100',
            'StartBit': '2',
            'Length': '4'
        }]
        mapped = self.extractor.map_and_clean(table)
        self.assertEqual(mapped[0]['Address'], '100_2_4')

        table2 = [{
            'Name': 'LenVar',
            'Address': '200',
            'Length': '10'
        }]
        mapped2 = self.extractor.map_and_clean(table2)
        self.assertEqual(mapped2[0]['Address'], '200_10')

if __name__ == '__main__':
    unittest.main()
