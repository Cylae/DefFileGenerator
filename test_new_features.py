import unittest
import csv
import os
import logging
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()
        self.generator = Generator()
        self.test_csv = "test_features.csv"

    def tearDown(self):
        if os.path.exists(self.test_csv):
            os.remove(self.test_csv)

    def test_address_offset_logic(self):
        # Test Simple Address Offset
        self.assertEqual(self.generator.apply_address_offset("30001", 10), "30011")
        # Test Compound Address Offset (Hex + String length)
        self.assertEqual(self.generator.apply_address_offset("0x30_20", 16), "64_20")
        # Test Compound Address Offset (Bits)
        self.assertEqual(self.generator.apply_address_offset("40001_0_1", 5), "40006_0_1")

    def test_tag_sanitization(self):
        # Multiple underscores and special characters
        rows = [{"Name": "Test  Variable!! @#", "Address": "30001"}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Tag'], "test_variable")

        # Strip underscores from start and end
        rows = [{"Name": "__Test Variable__", "Address": "30002"}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Tag'], "test_variable")

    def test_str_n_type_handling(self):
        # STR20 should be converted to STRING and address should be expanded if not already
        rows = [{"Name": "String test", "Address": "30050", "Type": "STR20"}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Info3'], "STRING")
        self.assertEqual(processed[0]['Info2'], "30050_20")

    def test_offset_application_during_mapping(self):
        # Mock extracted data
        raw_data = [{"Name": "Voltage", "Address": "30001"}]
        # map_and_clean should apply the offset
        mapped = self.extractor.map_and_clean(raw_data, address_offset=100)
        self.assertEqual(mapped[0]['Address'], "30101")

    def test_bits_type_overlap_prioritization(self):
        # Overlap check should allow BITS on the same address but warn for others
        rows = [
            {"Name": "Bit 0", "Address": "40001_0_1", "Type": "BITS"},
            {"Name": "Bit 1", "Address": "40001_1_1", "Type": "BITS"},
            {"Name": "Full Reg", "Address": "40001", "Type": "U16"}
        ]
        # This should log warnings but complete
        processed = self.generator.process_rows(rows)
        self.assertEqual(len(processed), 3)

if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    unittest.main()
