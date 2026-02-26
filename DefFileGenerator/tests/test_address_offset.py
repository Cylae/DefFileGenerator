import unittest
import logging
import io
import csv
from DefFileGenerator.def_gen import Generator

class TestAddressOffset(unittest.TestCase):
    def setUp(self):
        # Suppress logging warnings during tests
        logging.getLogger().setLevel(logging.ERROR)

    def test_positive_offset(self):
        gen = Generator(address_offset=1)
        rows = [
            {'Name': 'Test1', 'Address': '30001', 'Type': 'U16', 'RegisterType': 'Holding Register'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '30000')

    def test_negative_offset(self):
        gen = Generator(address_offset=-10)
        rows = [
            {'Name': 'Test1', 'Address': '30000', 'Type': 'U16', 'RegisterType': 'Holding Register'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '30010')

    def test_hex_address_with_offset(self):
        gen = Generator(address_offset=16)
        rows = [
            {'Name': 'TestHex', 'Address': '0x11', 'Type': 'U16', 'RegisterType': 'Holding Register'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '1') # 0x11 = 17, 17 - 16 = 1

    def test_negative_result_warning(self):
        gen = Generator(address_offset=100)
        rows = [
            {'Name': 'TestNeg', 'Address': '50', 'Type': 'U16', 'RegisterType': 'Holding Register'}
        ]
        with self.assertLogs(level='WARNING') as cm:
            processed = gen.process_rows(rows)
            self.assertEqual(processed[0]['Info2'], '-50')
            self.assertTrue(any("results in negative address -50" in output for output in cm.output))

    def test_string_address_with_offset(self):
        gen = Generator(address_offset=1)
        rows = [
            {'Name': 'TestStr', 'Address': '30001_10', 'Type': 'STRING', 'RegisterType': 'Holding Register'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '30000_10')

    def test_bits_address_with_offset(self):
        gen = Generator(address_offset=1)
        rows = [
            {'Name': 'TestBits', 'Address': '30001_0_1', 'Type': 'BITS', 'RegisterType': 'Holding Register'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '30000_0_1')

if __name__ == '__main__':
    unittest.main()
