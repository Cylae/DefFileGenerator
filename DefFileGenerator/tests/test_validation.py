import unittest
import os
import logging
import csv
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.gen = Generator()
        self.test_file = 'test_validation_output.csv'

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def test_modbus_range_validation(self):
        # Valid address
        self.assertTrue(self.gen.validate_address("100", "U16"))
        # Outside range (should still return True for format, but log warning)
        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.gen.validate_address("70000", "U16"))
        self.assertIn("outside standard Modbus range", cm.output[0])

    def test_validate_csv_success(self):
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])
            writer.writerow(['2', '3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4'])

        self.assertTrue(self.gen.validate_csv(self.test_file))

    def test_validate_csv_duplicate_tag(self):
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])
            writer.writerow(['2', '3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'A', '4'])

        with self.assertLogs(level='ERROR') as cm:
            self.assertFalse(self.gen.validate_csv(self.test_file))
        self.assertIn("Duplicate Tag 'tag1'", cm.output[0])

    def test_validate_csv_address_overlap(self):
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])
            writer.writerow(['2', '3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4'])

        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.gen.validate_csv(self.test_file))
        self.assertIn("Address overlap detected", cm.output[0])

    def test_validate_csv_insufficient_cols(self):
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U16']) # Only 4 cols

        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.gen.validate_csv(self.test_file))
        self.assertIn("Insufficient columns", cm.output[0])

if __name__ == '__main__':
    unittest.main()
