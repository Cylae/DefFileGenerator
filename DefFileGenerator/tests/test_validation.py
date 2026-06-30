import unittest
import os
import logging
import csv
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.valid_csv = "test_valid_def.csv"
        self.invalid_csv = "test_invalid_def.csv"

        # Create a valid definition CSV
        with open(self.valid_csv, 'w', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])
            writer.writerow(['2', '3', '30030_20', 'STRING', '', 'Var2', 'tag2', '1.0', '0.0', '', '4'])

    def tearDown(self):
        if os.path.exists(self.valid_csv):
            os.remove(self.valid_csv)
        if os.path.exists(self.invalid_csv):
            os.remove(self.invalid_csv)

    def test_validate_address_range(self):
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        self.assertTrue(self.generator.validate_address("0x0", "U16"))
        self.assertTrue(self.generator.validate_address("0xFFFF", "U16"))

        # Test out of range
        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(self.generator.validate_address("65536", "U16"))
            self.assertTrue(any("out of standard range" in m for m in log.output))

        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(self.generator.validate_address("-1", "U16"))
            self.assertTrue(any("out of standard range" in m for m in log.output))

    def test_validate_csv_success(self):
        self.assertTrue(self.generator.validate_csv(self.valid_csv))

    def test_validate_csv_failure(self):
        with open(self.invalid_csv, 'w', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            # Invalid type
            writer.writerow(['1', '3', '100', 'INVALID_TYPE', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])
            # Invalid address
            writer.writerow(['2', '3', '70000', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', '', '4'])

        self.assertFalse(self.generator.validate_csv(self.invalid_csv))

if __name__ == '__main__':
    unittest.main()
