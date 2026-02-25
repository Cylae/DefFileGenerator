import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestAddressOffset(unittest.TestCase):
    def setUp(self):
        # Disable logging for tests to keep output clean
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)

    def test_positive_offset(self):
        # Offset of 1 (convert 1-based to 0-based)
        gen = Generator(address_offset=1)
        rows = [
            {'Name': 'Test1', 'Address': '101', 'Type': 'U16'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '100')

    def test_hex_address_with_offset(self):
        gen = Generator(address_offset=16)
        rows = [
            {'Name': 'TestHex', 'Address': '0x20', 'Type': 'U16'}
        ]
        processed = gen.process_rows(rows)
        # 0x20 is 32. 32 - 16 = 16
        self.assertEqual(processed[0]['Info2'], '16')

    def test_negative_result_warning(self):
        gen = Generator(address_offset=100)
        rows = [
            {'Name': 'TestNeg', 'Address': '50', 'Type': 'U16'}
        ]
        # Should log a warning but still produce the result -50
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '-50')

    def test_composite_address_offset(self):
        # Offset should only apply to the base address part
        gen = Generator(address_offset=100)
        rows = [
            {'Name': 'TestString', 'Address': '1000_10', 'Type': 'STRING'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '900_10')

if __name__ == "__main__":
    unittest.main()
