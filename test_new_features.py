import unittest
import os
import csv
from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.gen = Generator()
        self.ext = Extractor()

    def test_address_offset_simple(self):
        self.assertEqual(self.gen.apply_address_offset("30001", 10), "30011")
        self.assertEqual(self.gen.apply_address_offset("0x10", 16), "32") # 16 + 16

    def test_address_offset_compound(self):
        self.assertEqual(self.gen.apply_address_offset("30001_20", 10), "30011_20")
        self.assertEqual(self.gen.apply_address_offset("40001_0_16", 5), "40006_0_16")

    def test_tag_sanitization(self):
        rows = [{'Name': 'Test Variable!!!', 'Address': '100', 'Type': 'U16'}]
        processed = self.gen.process_rows(rows)
        self.assertEqual(processed[0]['Tag'], 'test_variable')

        rows = [{'Name': '   Space   Test   ', 'Address': '101', 'Type': 'U16'}]
        processed = self.gen.process_rows(rows)
        self.assertEqual(processed[0]['Tag'], 'space_test')

    def test_str_n_handling(self):
        rows = [{'Name': 'StringTest', 'Address': '30050', 'Type': 'STR20'}]
        processed = self.gen.process_rows(rows)
        self.assertEqual(processed[0]['Info3'], 'STRING')
        self.assertEqual(processed[0]['Info2'], '30050_20')

    def test_extractor_complex_address(self):
        table = [[
            {'Name': 'Complex', 'Address': '1000', 'Length': '10', 'StartBit': '0'}
        ]]
        mapped = self.ext.map_and_clean(table)
        self.assertEqual(mapped[0]['Address'], '1000_0_10')

        table = [[
            {'Name': 'String', 'Address': '2000', 'Length': '20'}
        ]]
        mapped = self.ext.map_and_clean(table)
        self.assertEqual(mapped[0]['Address'], '2000_20')

    def test_extractor_offset_application(self):
        table = [[
            {'Name': 'Var1', 'Address': '1000', 'Type': 'U16'}
        ]]
        mapped = self.ext.map_and_clean(table, address_offset=100)
        self.assertEqual(mapped[0]['Address'], '1100')

if __name__ == "__main__":
    unittest.main()
