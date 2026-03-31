import unittest
import os
import csv
import io
import logging
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()
        self.generator = Generator()

    def test_address_offset_simple(self):
        # Test simple address offset
        res = self.generator.apply_address_offset("100", 10)
        self.assertEqual(res, "110")

    def test_address_offset_hex(self):
        # Test hex address with offset (hex normalized to decimal)
        res = self.generator.apply_address_offset("0x10", 10)
        self.assertEqual(res, "26")

    def test_address_offset_compound(self):
        # Test compound address (e.g., string or bit)
        res = self.generator.apply_address_offset("100_20", 10)
        self.assertEqual(res, "110_20")

        res = self.generator.apply_address_offset("100_5_1", 10)
        self.assertEqual(res, "110_5_1")

    def test_tag_sanitization(self):
        # Test tag generation from name with multiple underscores and special chars
        rows = [{"Name": "  Test   Variable!@#", "Address": "100"}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Tag'], "test_variable")

    def test_str_type_handling(self):
        # Test STR<n> type handling and address extension
        rows = [{"Name": "StringVar", "Address": "100", "Type": "STR20"}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Info3'], "STRING")
        self.assertEqual(processed[0]['Info2'], "100_20")

    def test_extraction_offset_application(self):
        # Test that map_and_clean applies the offset
        raw_tables = [[{"Name": "Test", "Address": "100"}]]
        mapped = self.extractor.map_and_clean(raw_tables, address_offset=50)
        self.assertEqual(mapped[0]['Address'], "150")

    def test_complex_address_construction(self):
        # Test building Address_StartBit_Length during extraction
        raw_tables = [[{"Name": "BitVar", "Address": "100", "StartBit": "2", "Length": "1"}]]
        mapped = self.extractor.map_and_clean(raw_tables)
        self.assertEqual(mapped[0]['Address'], "100_2_1")

        raw_tables = [[{"Name": "StrVar", "Address": "200", "Length": "10"}]]
        mapped = self.extractor.map_and_clean(raw_tables)
        self.assertEqual(mapped[0]['Address'], "200_10")

    def test_parse_numeric_european(self):
        # Test European decimal format
        self.assertEqual(self.generator._parse_numeric("1,23"), 1.23)
        self.assertEqual(self.generator._parse_numeric("1.23"), 1.23)
        self.assertEqual(self.generator._parse_numeric("1/10"), 0.1)

if __name__ == "__main__":
    logging.basicConfig(level=logging.ERROR)
    unittest.main()
