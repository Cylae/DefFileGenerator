import unittest
import os
import csv
import json
from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.extractor = Extractor()

    def test_apply_address_offset(self):
        # Simple integer
        self.assertEqual(self.generator.apply_address_offset("30001", 10), "30011")
        # Hex input should be normalized to decimal
        self.assertEqual(self.generator.apply_address_offset("0x10", 10), "26")
        # Compound address (Address_Length)
        self.assertEqual(self.generator.apply_address_offset("30001_20", 10), "30011_20")
        # Compound address (Address_StartBit_Length)
        self.assertEqual(self.generator.apply_address_offset("30001_0_1", 10), "30011_0_1")
        # Negative offset
        self.assertEqual(self.generator.apply_address_offset("30011", -10), "30001")

    def test_tag_sanitization(self):
        # The new RE_TAG_CLEAN should handle non-alphanumeric and underscores
        rows = [{'Name': 'Test @# Variable!', 'Address': '100', 'Type': 'U16'}]
        processed = self.generator.process_rows(rows)
        # "test_variable" (spaces and special chars replaced by _, then collapsed)
        self.assertEqual(processed[0]['Tag'], 'test_variable')

        # Test stripping and collapsing
        rows = [{'Name': '---Special---Tag---', 'Address': '101', 'Type': 'U16'}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Tag'], 'special_tag')

    def test_str_n_type_handling(self):
        rows = [{'Name': 'StrVar', 'Address': '30050', 'Type': 'STR20'}]
        # This is now handled in _normalize_type_and_address (called by process_rows)
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Info3'], 'STRING')
        self.assertEqual(processed[0]['Info2'], '30050_20')

    def test_mapping_offset_logic(self):
        # Test that map_and_clean uses apply_address_offset
        raw_tables = [[{'Name': 'Test', 'Offset': '30001', 'Type': 'U16'}]]
        mapped = self.extractor.map_and_clean(raw_tables, address_offset=10)
        self.assertEqual(mapped[0]['Address'], '30011')

    def test_double_offset_prevention(self):
        # Test that applying offset once in extractor and passing 0 to generator works correctly
        raw_tables = [[{'Name': 'Test', 'Address': '30001', 'Type': 'U16'}]]
        mapped = self.extractor.map_and_clean(raw_tables, address_offset=10) # 30011
        processed = self.generator.process_rows(mapped, address_offset=0) # Should still be 30011
        self.assertEqual(processed[0]['Info2'], '30011')

if __name__ == "__main__":
    unittest.main()
