import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.valid_csv = "test_valid_def.csv"
        self.invalid_csv = "test_invalid_def.csv"

        # Create a valid definition file
        with open(self.valid_csv, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '40001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])
            writer.writerow(['2', '3', '40002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4'])

    def tearDown(self):
        for f in [self.valid_csv, self.invalid_csv]:
            if os.path.exists(f):
                os.remove(f)

    def test_validate_csv_success(self):
        self.assertTrue(self.generator.validate_csv(self.valid_csv))

    def test_validate_csv_duplicate_tag(self):
        with open(self.invalid_csv, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '40001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])
            writer.writerow(['2', '3', '40002', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4'])

        with self.assertLogs(level='ERROR') as cm:
            self.assertFalse(self.generator.validate_csv(self.invalid_csv))
            self.assertTrue(any("Duplicate Tag 'tag1'" in m for m in cm.output))

    def test_validate_csv_out_of_range_address(self):
        with open(self.invalid_csv, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])

        with self.assertLogs(level='WARNING') as cm:
            self.assertFalse(self.generator.validate_csv(self.invalid_csv))
            self.assertTrue(any("outside standard range" in m for m in cm.output))

    def test_validate_csv_bad_format_address(self):
         with open(self.invalid_csv, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', 'not_an_address', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])

         with self.assertLogs(level='ERROR') as cm:
            self.assertFalse(self.generator.validate_csv(self.invalid_csv))
            self.assertTrue(any("Invalid Address 'not_an_address'" in m for m in cm.output))

    def test_address_range_validation(self):
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        with self.assertLogs(level='WARNING'):
            self.assertFalse(self.generator.validate_address("65536", "U16"))
        with self.assertLogs(level='WARNING'):
            self.assertFalse(self.generator.validate_address("-1", "U16"))

    def test_str_n_address_validation(self):
        self.assertTrue(self.generator.validate_address("100_20", "STR20"))
        self.assertTrue(self.generator.validate_address("0x64_20", "STR20"))
        with self.assertLogs(level='WARNING'):
            self.assertFalse(self.generator.validate_address("70000_20", "STR20"))

if __name__ == "__main__":
    unittest.main()
