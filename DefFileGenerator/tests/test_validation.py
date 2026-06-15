import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        logging.basicConfig(level=logging.INFO)

    def test_validate_address_range(self):
        # Standard Modbus range
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))

        # Out of range (should log warning but return True for flexibility)
        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_address("65536", "U16"))
            self.assertIn("outside standard Modbus range", cm.output[0])

    def test_validate_address_hex(self):
        self.assertTrue(self.generator.validate_address("0xFFFF", "U16"))
        self.assertTrue(self.generator.validate_address("FFFFh", "U16"))

        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_address("0x10000", "U16"))
            self.assertIn("outside standard Modbus range", cm.output[0])

    def test_validate_address_compound(self):
        # BITS
        self.assertTrue(self.generator.validate_address("100_0_1", "BITS"))
        # STRING
        self.assertTrue(self.generator.validate_address("200_10", "STRING"))

        # Invalid format
        self.assertFalse(self.generator.validate_address("invalid", "U16"))
        self.assertFalse(self.generator.validate_address("100_0", "BITS")) # Missing len
        self.assertFalse(self.generator.validate_address("200", "STRING")) # Missing len

if __name__ == '__main__':
    unittest.main()
