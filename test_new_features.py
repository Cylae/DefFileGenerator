#!/usr/bin/env python3
import unittest
import os
import sys
import csv
from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.extractor = Extractor()

    def test_address_offset(self):
        # Simple address
        self.assertEqual(self.generator.apply_address_offset("30001", 10), "30011")
        self.assertEqual(self.generator.apply_address_offset("0x10", 10), "26") # 0x10 = 16

        # Compound address
        self.assertEqual(self.generator.apply_address_offset("30001_20", 10), "30011_20")
        self.assertEqual(self.generator.apply_address_offset("30001_16_0", 10), "30011_16_0")

    def test_tag_sanitization(self):
        rows = [{'Name': 'Test Variable @ 123', 'Address': '1', 'Type': 'U16'}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Tag'], 'test_variable_123')

        rows = [{'Name': '   Spaces   Around   ', 'Address': '2', 'Type': 'U16'}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Tag'], 'spaces_around')

    def test_compound_address_generation(self):
        # Test extraction logic for compound addresses
        table = [{
            'Name': 'String Var',
            'Address': '30050',
            'Length': '20',
            'Type': 'STRING'
        }]
        mapped = self.extractor.map_and_clean([table])
        self.assertEqual(mapped[0]['Address'], '30050_20')

        table = [{
            'Name': 'Bit Var',
            'Address': '40001',
            'Length': '1',
            'StartBit': '5',
            'Type': 'BITS'
        }]
        mapped = self.extractor.map_and_clean([table])
        self.assertEqual(mapped[0]['Address'], '40001_1_5')

    def test_convenience_type_normalization(self):
        # STR<n> should be normalized to STRING and address should be expanded
        rows = [{'Name': 'Convenience String', 'Address': '30001', 'Type': 'STR20'}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Info3'], 'STRING')
        self.assertEqual(processed[0]['Info2'], '30001_20')

    def test_optimized_overlap_check(self):
        # Ensure no crash and correct detection for large number of registers
        rows = []
        for i in range(100):
            rows.append({'Name': f'Var{i}', 'Address': str(30000 + i), 'Type': 'U16'})

        # Overlapping register
        rows.append({'Name': 'Overlapper', 'Address': '30050', 'Type': 'U16'})

        with self.assertLogs(level='WARNING') as cm:
            self.generator.process_rows(rows)
            self.assertTrue(any("Address overlap detected" in log for log in cm.output))

if __name__ == "__main__":
    unittest.main()
