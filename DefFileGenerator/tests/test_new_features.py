import unittest
import os
import csv
import logging
import tempfile
from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()

    def test_address_offset(self):
        # Test with offset 1 (convert 1-based to 0-based)
        gen = Generator(address_offset=1)
        rows = [
            {'Name': 'Test1', 'Address': '1', 'Type': 'U16'},
            {'Name': 'Test2', 'Address': '101', 'Type': 'U16'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '0')
        self.assertEqual(processed[1]['Info2'], '100')

    def test_address_offset_hex(self):
        # Test with hex and offset
        gen = Generator(address_offset=16)
        rows = [
            {'Name': 'Test1', 'Address': '0x10', 'Type': 'U16'},
            {'Name': 'Test2', 'Address': '20h', 'Type': 'U16'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '0') # 0x10 = 16, 16-16=0
        self.assertEqual(processed[1]['Info2'], '16') # 20h = 32, 32-16=16

    def test_negative_address_warning(self):
        gen = Generator(address_offset=100)
        rows = [
            {'Name': 'Test1', 'Address': '10', 'Type': 'U16'}
        ]
        with self.assertLogs(level='WARNING') as cm:
            processed = gen.process_rows(rows)
            self.assertTrue(any("results in negative address -90" in output for output in cm.output))
        self.assertEqual(processed[0]['Info2'], '-90')

    def test_xml_extraction(self):
        # Create a dummy XML
        xml_content = """<?xml version='1.0' encoding='utf-8'?>
<root>
  <row>
    <Name>Voltage</Name>
    <Address>30001</Address>
    <Type>U16</Type>
  </row>
  <row>
    <Name>Current</Name>
    <Address>30002</Address>
    <Type>U16</Type>
  </row>
</root>
"""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.xml', delete=False) as tf:
            tf.write(xml_content)
            temp_xml = tf.name

        try:
            tables = self.extractor.extract_from_xml(temp_xml)
            self.assertEqual(len(tables), 1)
            self.assertEqual(len(tables[0]), 2)
            self.assertEqual(tables[0][0]['Name'], 'Voltage')
        finally:
            if os.path.exists(temp_xml):
                os.remove(temp_xml)

    def test_csv_extraction(self):
        # Create a dummy CSV with ; delimiter
        csv_content = "Name;Address;Type\nVoltage;30001;U16\nCurrent;30002;U16\n"
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as tf:
            tf.write(csv_content)
            temp_csv = tf.name

        try:
            tables = self.extractor.extract_from_csv(temp_csv)
            self.assertEqual(len(tables), 1)
            self.assertEqual(len(tables[0]), 2)
            self.assertEqual(tables[0][0]['Name'], 'Voltage')
        finally:
            if os.path.exists(temp_csv):
                os.remove(temp_csv)

    def test_normalize_address_val_messy(self):
        gen = Generator()
        # Test messy strings from PDF extraction
        self.assertEqual(gen.normalize_address_val("Reg: 0x10 (Holding)"), "16")
        self.assertEqual(gen.normalize_address_val("Address 40001"), "40001")
        self.assertEqual(gen.normalize_address_val("10,000"), "10000")

if __name__ == '__main__':
    unittest.main()
