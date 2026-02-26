import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestAddressOffset(unittest.TestCase):
    def test_address_offset(self):
        # Test with offset 1
        generator = Generator(address_offset=1)
        rows = [{
            'Name': 'Test Var',
            'Tag': 'test_tag',
            'RegisterType': 'Holding Register',
            'Address': '30001',
            'Type': 'U16',
            'Factor': '1',
            'Offset': '0',
            'Unit': 'V',
            'Action': '4',
            'ScaleFactor': '0'
        }]
        processed = generator.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '30000')

    def test_address_offset_string(self):
        # Test with offset 10 and a string type
        generator = Generator(address_offset=10)
        rows = [{
            'Name': 'Test Str',
            'Tag': 'str_tag',
            'RegisterType': 'Holding Register',
            'Address': '30050_20',
            'Type': 'STRING',
            'Factor': '1',
            'Offset': '0',
            'Unit': '',
            'Action': '4',
            'ScaleFactor': '0'
        }]
        processed = generator.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '30040_20')

    def test_negative_address_warning(self):
        generator = Generator(address_offset=100)
        rows = [{
            'Name': 'Test Var',
            'Tag': 'test_tag',
            'RegisterType': 'Holding Register',
            'Address': '50',
            'Type': 'U16',
            'Factor': '1',
            'Offset': '0',
            'Unit': 'V',
            'Action': '4',
            'ScaleFactor': '0'
        }]
        with self.assertLogs(level='WARNING') as log:
            processed = generator.process_rows(rows)
            self.assertEqual(processed[0]['Info2'], '-50')
            self.assertTrue(any("results in negative address -50" in m for m in log.output))

if __name__ == '__main__':
    unittest.main()
