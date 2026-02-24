import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestAddressOffset(unittest.TestCase):
    def test_address_offset_subtraction(self):
        # Test basic offset subtraction
        generator = Generator(address_offset=10)
        rows = [
            {'Name': 'Test1', 'RegisterType': 'Holding Register', 'Address': '100', 'Type': 'U16'}
        ]
        processed = generator.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '90')

    def test_hex_address_with_offset(self):
        # Test hex address with offset
        generator = Generator(address_offset=10)
        rows = [
            {'Name': 'Test1', 'RegisterType': 'Holding Register', 'Address': '0x64', 'Type': 'U16'}
        ]
        # 0x64 is 100. 100 - 10 = 90
        processed = generator.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '90')

    def test_composite_address_with_offset(self):
        # Test Address_Length format (STRING)
        generator = Generator(address_offset=100)
        rows = [
            {'Name': 'TestString', 'RegisterType': 'Holding Register', 'Address': '1000_20', 'Type': 'STRING'}
        ]
        processed = generator.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '900_20')

    def test_negative_address_warning(self):
        # Test that negative address triggers a warning
        generator = Generator(address_offset=100)
        rows = [
            {'Name': 'TestNeg', 'RegisterType': 'Holding Register', 'Address': '50', 'Type': 'U16'}
        ]
        with self.assertLogs(level='WARNING') as cm:
            processed = generator.process_rows(rows)
            self.assertEqual(processed[0]['Info2'], '-50')
            self.assertTrue(any("results in negative address -50" in output for output in cm.output))

if __name__ == '__main__':
    unittest.main()
