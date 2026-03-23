import unittest
import os
import csv
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()
        self.generator = Generator()
        self.csv_file = "test_features.csv"

    def tearDown(self):
        if os.path.exists(self.csv_file):
            os.remove(self.csv_file)

    def test_address_offset_simple(self):
        addr = "100"
        offset = 10
        result = self.generator.apply_address_offset(addr, offset)
        self.assertEqual(result, "110")

    def test_address_offset_hex(self):
        addr = "0x10"
        offset = 10
        result = self.generator.apply_address_offset(addr, offset)
        self.assertEqual(result, "26") # 16 + 10 = 26

    def test_address_offset_compound_string(self):
        addr = "100_20" # Addr_Len
        offset = 10
        result = self.generator.apply_address_offset(addr, offset)
        self.assertEqual(result, "110_20")

    def test_address_offset_compound_bits(self):
        addr = "100_0_1" # Addr_Bit_Len
        offset = 10
        result = self.generator.apply_address_offset(addr, offset)
        self.assertEqual(result, "110_0_1")

    def test_compound_address_from_columns(self):
        # Test building Addr_Len from Address and Length columns
        tables = [[
            {"Address": "100", "Name": "Var1", "Length": "20", "Type": "STRING"}
        ]]
        mapped = self.extractor.map_and_clean(tables)
        self.assertEqual(mapped[0]["Address"], "100_20")

        # Test building Addr_Bit_Len from Address, StartBit and Length columns
        tables = [[
            {"Address": "200", "Name": "Var2", "StartBit": "0", "Length": "1", "Type": "BITS"}
        ]]
        mapped = self.extractor.map_and_clean(tables)
        self.assertEqual(mapped[0]["Address"], "200_0_1")

    def test_extraction_with_offset(self):
        tables = [[
            {"Address": "100", "Name": "Var1", "Type": "U16"}
        ]]
        mapped = self.extractor.map_and_clean(tables, address_offset=50)
        self.assertEqual(mapped[0]["Address"], "150")

    def test_tag_cleaning(self):
        # Tags should only contain a-z, 0-9, and _
        rows = [{"Name": "Test!@#$%^&*()Variable", "Address": "100", "Type": "U16"}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]["Tag"], "test_variable")

if __name__ == "__main__":
    unittest.main()
