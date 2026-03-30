#!/usr/bin/env python3
import unittest
import logging
from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.extractor = Extractor()

    def test_address_offset_simple(self):
        addr = "30001"
        offset = 10
        self.assertEqual(self.generator.apply_address_offset(addr, offset), "30011")

    def test_address_offset_compound(self):
        # STR20 address format: Address_Length
        addr = "30001_20"
        offset = 10
        self.assertEqual(self.generator.apply_address_offset(addr, offset), "30011_20")

    def test_address_offset_hex(self):
        # Hex address should be normalized to decimal before/during offset
        addr = "0x10" # 16 decimal
        offset = 10
        self.assertEqual(self.generator.apply_address_offset(addr, offset), "26")

    def test_tag_sanitization(self):
        name = "Test Variable @#$ With   Spaces"
        tag = self.generator._process_name_and_tag(name, None, {}, 2)
        # Should collapse underscores and strip them from ends
        self.assertEqual(tag, "test_variable_with_spaces")

    def test_str_n_type_handling(self):
        rows = [{'Name': 'TestStr', 'Address': '30050', 'Type': 'STR20', 'RegisterType': 'Holding'}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Info3'], 'STRING')
        self.assertEqual(processed[0]['Info2'], '30050_20')

    def test_extractor_complex_address(self):
        # Test building Address_StartBit_Length in map_and_clean
        tables = [[{'Name': 'BitVar', 'Address': '100', 'Length': '1', 'StartBit': '5', 'RegisterType': 'Holding'}]]
        mapped = self.extractor.map_and_clean(tables)
        self.assertEqual(mapped[0]['Address'], '100_5_1')

    def test_extractor_address_offset_application(self):
        tables = [[{'Name': 'Var1', 'Address': '100', 'RegisterType': 'Holding'}]]
        mapped = self.extractor.map_and_clean(tables, address_offset=1000)
        self.assertEqual(mapped[0]['Address'], '1100')

if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    unittest.main()
