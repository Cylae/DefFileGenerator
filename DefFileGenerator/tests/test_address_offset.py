import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestAddressOffset(unittest.TestCase):
    def test_address_offset(self):
        generator = Generator(address_offset=10)
        rows = [{
            'Name': 'Test Var',
            'Address': '30010',
            'Type': 'U16'
        }]
        processed = generator.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '30000')

    def test_negative_address_warning(self):
        generator = Generator(address_offset=100)
        rows = [{
            'Name': 'Test Var',
            'Address': '50',
            'Type': 'U16'
        }]
        with self.assertLogs(level='WARNING') as log:
            processed = generator.process_rows(rows)
            self.assertEqual(processed[0]['Info2'], '-50')
            self.assertTrue(any("results in negative address -50" in m for m in log.output))

    def test_hex_address_offset(self):
        generator = Generator(address_offset=1)
        rows = [{
            'Name': 'Test Var',
            'Address': '0x11', # 17 decimal
            'Type': 'U16'
        }]
        processed = generator.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '16')

    def test_composite_address_offset(self):
        generator = Generator(address_offset=10)
        rows = [{
            'Name': 'String Var',
            'Address': '30010_10',
            'Type': 'STRING'
        }]
        processed = generator.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '30000_10')

if __name__ == '__main__':
    unittest.main()
