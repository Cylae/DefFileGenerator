import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_validate_address_range(self):
        with self.assertLogs(level='WARNING') as log:
            # Valid range
            self.assertTrue(self.generator.validate_address('0', 'U16'))
            self.assertTrue(self.generator.validate_address('65535', 'U16'))

            # Outside range (should still return True but log warning)
            self.assertTrue(self.generator.validate_address('65536', 'U16'))
            self.assertTrue(any("outside standard Modbus range" in m for m in log.output))

            self.assertTrue(self.generator.validate_address('-1', 'U16'))
            self.assertTrue(any("outside standard Modbus range" in m for m in log.output))

    def test_validate_address_hex_range(self):
        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(self.generator.validate_address('0xFFFF', 'U16'))
            self.assertTrue(self.generator.validate_address('0x10000', 'U16'))
            self.assertTrue(any("outside standard Modbus range" in m for m in log.output))

if __name__ == '__main__':
    unittest.main()
