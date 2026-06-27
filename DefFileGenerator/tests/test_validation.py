import unittest
import os
import csv
import tempfile
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_dir = tempfile.TemporaryDirectory()
        logging.basicConfig(level=logging.ERROR)

    def tearDown(self):
        self.test_dir.cleanup()

    def test_validate_address_range(self):
        # Valid range
        self.assertTrue(Generator.validate_address("0", "U16"))
        self.assertTrue(Generator.validate_address("65535", "U16"))
        self.assertTrue(Generator.validate_address("0x0", "U16"))
        self.assertTrue(Generator.validate_address("0xFFFF", "U16"))

        # Out of range
        self.assertFalse(Generator.validate_address("65536", "U16"))
        self.assertFalse(Generator.validate_address("-1", "U16"))
        self.assertFalse(Generator.validate_address("0x10000", "U16"))

    def test_validate_compound_address_range(self):
        # Valid compound
        self.assertTrue(Generator.validate_address("100_20", "STRING"))
        self.assertTrue(Generator.validate_address("0x100_0_1", "BITS"))

        # Invalid compound range
        self.assertFalse(Generator.validate_address("66000_20", "STRING"))

    def test_validate_csv_format(self):
        csv_path = os.path.join(self.test_dir.name, "test_valid.csv")
        with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Test', 'Test', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U16', '', 'Name', 'tag', '1.0', '0.0', 'V', '4'])

        self.assertTrue(Generator.validate_csv(csv_path))

    def test_validate_csv_invalid_address(self):
        csv_path = os.path.join(self.test_dir.name, "test_invalid_addr.csv")
        with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Test', 'Test', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '70000', 'U16', '', 'Name', 'tag', '1.0', '0.0', 'V', '4'])

        self.assertFalse(Generator.validate_csv(csv_path))

    def test_validate_csv_insufficient_columns(self):
        csv_path = os.path.join(self.test_dir.name, "test_short_row.csv")
        with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Test', 'Test', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U16']) # Only 4 columns

        # Should skip the row and return True if no other errors, but let's check behavior
        # Current implementation returns True if no INVALID addresses found.
        self.assertTrue(Generator.validate_csv(csv_path))

if __name__ == '__main__':
    unittest.main()
