#!/usr/bin/env python3
import unittest
from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.extractor = Extractor()

    def test_apply_address_offset(self):
        # Simple address
        self.assertEqual(self.generator.apply_address_offset("100", 50), "150")
        # Hex address
        self.assertEqual(self.generator.apply_address_offset("0x64", 50), "150")
        # Compound address (Address_Length)
        self.assertEqual(self.generator.apply_address_offset("100_10", 50), "150_10")
        # Compound address (Address_StartBit_Length)
        self.assertEqual(self.generator.apply_address_offset("100_0_16", 50), "150_0_16")
        # Zero offset still normalizes
        self.assertEqual(self.generator.apply_address_offset("0x64", 0), "100")

    def test_tag_sanitization(self):
        rows = [
            {"Name": "Test-Variable!", "Address": "100", "Type": "U16"},
            {"Name": "Test Variable", "Address": "101", "Type": "U16"},
            {"Name": "!!@@##", "Address": "102", "Type": "U16"}
        ]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]["Tag"], "test_variable")
        self.assertEqual(processed[1]["Tag"], "test_variable_1")
        self.assertEqual(processed[2]["Tag"], "var")

    def test_str_n_handling(self):
        rows = [
            {"Name": "String1", "Address": "100", "Type": "STR20"},
            {"Name": "String2", "Address": "200_10", "Type": "STR10"}
        ]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]["Info3"], "STRING")
        self.assertEqual(processed[0]["Info2"], "100_20")
        self.assertEqual(processed[1]["Info3"], "STRING")
        self.assertEqual(processed[1]["Info2"], "200_10")

    def test_offset_in_mapping(self):
        raw_tables = [[
            {"Addr": "100", "Label": "Var1", "Format": "U16"},
            {"Addr": "0x200", "Label": "Var2", "Format": "U16"}
        ]]
        self.extractor.mapping = {"Address": "Addr", "Name": "Label", "Type": "Format"}
        mapped = self.extractor.map_and_clean(raw_tables, address_offset=100)
        self.assertEqual(mapped[0]["Address"], "200")
        self.assertEqual(mapped[1]["Address"], "612") # 0x200 = 512, + 100 = 612

    def test_negative_address_warning(self):
        rows = [
            {"Name": "NegVar", "Address": "50", "Type": "U16"}
        ]
        import logging
        with self.assertLogs(level='WARNING') as cm:
            self.generator.process_rows(rows, address_offset=-100)
            self.assertTrue(any("results in negative address -50" in output for output in cm.output))

if __name__ == "__main__":
    unittest.main()
