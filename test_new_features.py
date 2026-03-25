import unittest
import csv
import os
import sys
from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.extractor = Extractor()

    def test_address_offset_simple(self):
        addr = "100"
        offset = 50
        result = self.generator.apply_address_offset(addr, offset)
        self.assertEqual(result, "150")

    def test_address_offset_compound_string(self):
        addr = "30001_20"
        offset = 10
        result = self.generator.apply_address_offset(addr, offset)
        self.assertEqual(result, "30011_20")

    def test_address_offset_compound_bits(self):
        addr = "40001_5_2"
        offset = 100
        result = self.generator.apply_address_offset(addr, offset)
        self.assertEqual(result, "40101_5_2")

    def test_address_offset_hex(self):
        addr = "0x10"
        offset = 10
        # normalize_address_val converts 0x10 to 16
        # 16 + 10 = 26
        result = self.generator.apply_address_offset(addr, offset)
        self.assertEqual(result, "26")

    def test_tag_sanitization(self):
        # Name: "Battery (Power) % Level"
        # Tag should become: "battery_power_level"
        rows = [{"Name": "Battery (Power) % Level", "Address": "100", "Type": "U16"}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]["Tag"], "battery_power_level")

    def test_str_n_type_handling(self):
        rows = [{"Name": "ModelName", "Address": "30050", "Type": "STR20"}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]["Info3"], "STRING")
        self.assertEqual(processed[0]["Info2"], "30050_20")

    def test_mapping_normalization(self):
        raw_tables = [[{"name": "test", "address": "100", "type": "u16"}]]
        mapped = self.extractor.map_and_clean(raw_tables, address_offset=10)
        self.assertEqual(mapped[0]["Address"], "110")
        self.assertEqual(mapped[0]["Type"], "U16")

if __name__ == "__main__":
    unittest.main()
