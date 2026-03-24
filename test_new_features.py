import unittest
import os
import csv
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.extractor = Extractor()

    def test_address_offset_simple(self):
        # Simple address
        addr = "100"
        offset = 10
        result = self.generator.apply_address_offset(addr, offset)
        self.assertEqual(result, "110")

    def test_address_offset_compound(self):
        # Compound address (Address_Length)
        addr = "100_20"
        offset = 10
        result = self.generator.apply_address_offset(addr, offset)
        self.assertEqual(result, "110_20")

    def test_address_offset_bits(self):
        # Compound address (Address_Length_StartBit)
        addr = "100_1_5"
        offset = 10
        result = self.generator.apply_address_offset(addr, offset)
        self.assertEqual(result, "110_1_5")

    def test_compound_address_generation(self):
        # Test extraction logic for compound addresses
        table = [{
            'Name': 'Test Var',
            'Address': '1000',
            'Length': '2',
            'StartBit': '0'
        }]
        mapped = self.extractor.map_and_clean(table)
        self.assertEqual(mapped[0]['Address'], '1000_2_0')

        table2 = [{
            'Name': 'Test Var 2',
            'Address': '2000',
            'Length': '10'
        }]
        mapped2 = self.extractor.map_and_clean(table2)
        self.assertEqual(mapped2[0]['Address'], '2000_10')

        table3 = [{
            'Name': 'Test Var 3',
            'Address': '3000',
            'StartBit': '5'
        }]
        mapped3 = self.extractor.map_and_clean(table3)
        self.assertEqual(mapped3[0]['Address'], '3000_1_5')

    def test_tag_sanitization(self):
        rows = [{'Name': 'AC Power (W)', 'Address': '100', 'Type': 'U16'}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Tag'], 'ac_power_w')

        rows2 = [{'Name': 'DC-Voltage!!!', 'Address': '101', 'Type': 'U16'}]
        processed2 = self.generator.process_rows(rows2)
        self.assertEqual(processed2[0]['Tag'], 'dc_voltage')

    def test_address_overlap_optimized(self):
        # Test overlap detection with O(1) logic
        rows = [
            {'Name': 'Var 1', 'Address': '100', 'Type': 'U32'}, # Uses 100, 101
            {'Name': 'Var 2', 'Address': '101', 'Type': 'U16'}  # Should overlap
        ]
        with self.assertLogs(level='WARNING') as cm:
            self.generator.process_rows(rows)
            self.assertTrue(any("Address overlap detected" in log for log in cm.output))

    def test_negative_address_offset(self):
        rows = [{'Name': 'Var 1', 'Address': '10', 'Type': 'U16'}]
        with self.assertLogs(level='WARNING') as cm:
            self.generator.process_rows(rows, address_offset=-20)
            self.assertTrue(any("negative address -10" in log for log in cm.output))

if __name__ == "__main__":
    unittest.main()
