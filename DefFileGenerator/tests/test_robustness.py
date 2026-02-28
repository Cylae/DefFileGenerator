import unittest
from DefFileGenerator.def_gen import Generator
import logging

class TestRobustness(unittest.TestCase):
    def setUp(self):
        # Disable logging for tests to keep output clean
        logging.disable(logging.CRITICAL)

    def test_address_offset(self):
        # Test basic offset
        gen = Generator(address_offset=1)
        rows = [{'Name': 'Test', 'Address': '10', 'Type': 'U16', 'RegisterType': 'Holding Register'}]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '9')

        # Test hex address with offset
        gen = Generator(address_offset=0x10)
        rows = [{'Name': 'Test', 'Address': '0x20', 'Type': 'U16', 'RegisterType': 'Holding Register'}]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '16') # 32 - 16 = 16

    def test_negative_address_offset(self):
        # Test negative result (should log warning but still calculate)
        gen = Generator(address_offset=100)
        rows = [{'Name': 'Test', 'Address': '10', 'Type': 'U16', 'RegisterType': 'Holding Register'}]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '-90')

    def test_composite_address_with_offset(self):
        # Test STR type (Address_Length)
        gen = Generator(address_offset=10)
        rows = [{'Name': 'Test', 'Address': '100_10', 'Type': 'STRING', 'RegisterType': 'Holding Register'}]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '90_10')

        # Test BITS type (Address_StartBit_NbBits)
        gen = Generator(address_offset=10)
        rows = [{'Name': 'Test', 'Address': '100_0_1', 'Type': 'BITS', 'RegisterType': 'Holding Register'}]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '90_0_1')

    def test_hex_composite_address(self):
        gen = Generator()
        # Hex address with length
        rows = [{'Name': 'Test', 'Address': '0x10_10', 'Type': 'STRING', 'RegisterType': 'Holding Register'}]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '16_10')

        # Hex address with bits
        rows = [{'Name': 'Test', 'Address': '0x10_0_1', 'Type': 'BITS', 'RegisterType': 'Holding Register'}]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '16_0_1')

    def test_numeric_validation_fallbacks(self):
        gen = Generator()
        # Invalid Factor, Offset, ScaleFactor
        rows = [{
            'Name': 'Test',
            'Address': '10',
            'Type': 'U16',
            'Factor': 'invalid',
            'Offset': 'nan',
            'ScaleFactor': 'foo'
        }]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['CoefA'], '1.000000') # Default factor 1.0, scale 0
        self.assertEqual(processed[0]['CoefB'], '0.000000') # Default offset 0.0

    def test_fractional_factor(self):
        gen = Generator()
        # Test 1/100 format
        rows = [{'Name': 'Test', 'Address': '10', 'Type': 'U16', 'Factor': '1/100'}]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['CoefA'], '0.010000')

    def tearDown(self):
        logging.disable(logging.NOTSET)

if __name__ == '__main__':
    unittest.main()
