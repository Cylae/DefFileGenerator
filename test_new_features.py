import os
import csv
import json
import unittest
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.test_csv = 'test_registers.csv'
        with open(self.test_csv, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Parameter', 'Addr', 'Data Type', 'Unit', 'Ratio', 'Bit Offset', 'Length'])
            writer.writerow(['Voltage', '40001', 'uint16', 'V', '0.1', '', ''])
            writer.writerow(['Current', '0x9C42', 'int16', 'A', '1/100', '', ''])
            writer.writerow(['Status Bits', '30005', 'bits', '', '', '0', '2'])
            writer.writerow(['Device Name', '30010', 'STR10', '', '', '', ''])

    def tearDown(self):
        if os.path.exists(self.test_csv):
            os.remove(self.test_csv)
        if os.path.exists('test_output.csv'):
            os.remove('test_output.csv')

    def test_extraction_and_offset(self):
        extractor = Extractor()
        raw = extractor.extract_from_csv(self.test_csv)
        # Apply offset 10 during extraction
        mapped = extractor.map_and_clean(raw, address_offset=10)

        # Verify addresses
        # 40001 + 10 = 40011
        # 0x9C42 (40002) + 10 = 40012
        # 30005 + 10 = 30015 (with bits)
        # 30010 + 10 = 30020

        self.assertEqual(mapped[0]['Address'], '40011')
        self.assertEqual(mapped[1]['Address'], '40012')
        self.assertEqual(mapped[2]['Address'], '30015_0_2')
        self.assertEqual(mapped[3]['Address'], '30020')

    def test_generation_logic(self):
        generator = Generator()
        rows = [
            {'Name': 'Voltage', 'Address': '40011', 'Type': 'U16', 'Factor': '0.1', 'Unit': 'V'},
            {'Name': 'Current', 'Address': '40012', 'Type': 'I16', 'Factor': '0,01', 'Unit': 'A'}, # European decimal
            {'Name': 'Status', 'Address': '30015_0_2', 'Type': 'BITS', 'Unit': ''},
            {'Name': 'Name', 'Address': '30020_10', 'Type': 'STRING', 'Unit': ''}
        ]
        processed = generator.process_rows(rows, address_offset=0)

        self.assertEqual(processed[0]['CoefA'], '0.100000')
        self.assertEqual(processed[1]['CoefA'], '0.010000')
        self.assertEqual(processed[2]['Info2'], '30015_0_2')
        self.assertEqual(processed[3]['Info3'], 'STRING')

    def test_european_and_fractions(self):
        generator = Generator()
        self.assertEqual(generator._parse_numeric('0,5'), 0.5)
        self.assertEqual(generator._parse_numeric('1/10'), 0.1)

if __name__ == '__main__':
    unittest.main()
