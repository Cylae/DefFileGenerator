import unittest
from DefFileGenerator.extractor import Extractor

class TestRobustness(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()

    def test_multi_pass_matching(self):
        # Pass A: Exact/Normalized should win over Pass B: Fuzzy
        raw_data = [[
            {"Name": "ExactName", "Parameter": "FuzzyName", "Address": "100"}
        ]]
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Name"], "ExactName")

    def test_bits_address_defaulting(self):
        raw_data = [[
            {"Name": "BitVar", "Address": "100", "Type": "BITS", "StartBit": "3", "Length": "2"}
        ]]
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "100_3_2")

    def test_string_address_defaulting(self):
        raw_data = [[
            {"Name": "StrVar", "Address": "200", "Type": "STRING", "Length": "10"}
        ]]
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "200_10")

    def test_thousands_separator_normalization(self):
        raw_data = [[
            {"Name": "SepVar", "Address": "1,234", "Type": "U16"}
        ]]
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "1234")

if __name__ == "__main__":
    unittest.main()
