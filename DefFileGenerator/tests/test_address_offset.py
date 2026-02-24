import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestAddressOffset(unittest.TestCase):
    def test_address_offset_basic(self):
        # Subtract 1 from addresses
        generator = Generator(address_offset=1)
        rows = [{
            'Name': 'Test Var',
            'Address': '30001',
            'Type': 'U16',
            'RegisterType': 'Holding Register'
        }]
        processed = generator.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '30000')

    def test_address_offset_hex(self):
        # Subtract 16 from 0x20 (32) -> 16 (0x10)
        generator = Generator(address_offset=16)
        rows = [{
            'Name': 'Test Hex',
            'Address': '0x20',
            'Type': 'U16',
            'RegisterType': 'Holding Register'
        }]
        processed = generator.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '16')

    def test_address_offset_string(self):
        # Subtract 1 from 30001_10 -> 30000_10
        generator = Generator(address_offset=1)
        rows = [{
            'Name': 'Test String',
            'Address': '30001_10',
            'Type': 'STRING',
            'RegisterType': 'Holding Register'
        }]
        processed = generator.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '30000_10')

    def test_address_offset_bits(self):
        # Subtract 1 from 30001_0_1 -> 30000_0_1
        generator = Generator(address_offset=1)
        rows = [{
            'Name': 'Test Bits',
            'Address': '30001_0_1',
            'Type': 'BITS',
            'RegisterType': 'Holding Register'
        }]
        processed = generator.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '30000_0_1')

    def test_negative_address_warning(self):
        # Subtract 10 from 5 -> -5. Should log a warning.
        generator = Generator(address_offset=10)
        rows = [{
            'Name': 'Test Neg',
            'Address': '5',
            'Type': 'U16',
            'RegisterType': 'Holding Register'
        }]
        with self.assertLogs(level='WARNING') as log:
            processed = generator.process_rows(rows)
            self.assertEqual(processed[0]['Info2'], '-5')
            self.assertTrue(any("results in negative address -5" in m for m in log.output))

if __name__ == '__main__':
    unittest.main()
