import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_file = "test_validate.csv"

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def write_csv(self, rows):
        with open(self.test_file, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            for row in rows:
                writer.writerow(row)

    def test_valid_file(self):
        self.write_csv([
            ['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''],
            ['1', '3', '100', 'U16', '', 'Name1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['2', '3', '101', 'U16', '', 'Name2', 'tag2', '1.0', '0.0', 'V', '4']
        ])
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_duplicate_tag(self):
        self.write_csv([
            ['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''],
            ['1', '3', '100', 'U16', '', 'Name1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['2', '3', '101', 'U16', '', 'Name2', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        with self.assertLogs(level='ERROR') as cm:
            self.assertFalse(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Duplicate Tag 'tag1'" in msg for msg in cm.output))

    def test_address_overlap(self):
        self.write_csv([
            ['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''],
            ['1', '3', '100', 'U32', '', 'Name1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['2', '3', '101', 'U16', '', 'Name2', 'tag2', '1.0', '0.0', 'V', '4']
        ])
        with self.assertLogs(level='WARNING') as cm:
            # Overlap is currently a warning, so validation might still return True if no other errors
            self.assertTrue(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Address overlap detected for 'Name2' at 101" in msg for msg in cm.output))

    def test_invalid_header(self):
        self.write_csv([
            ['too', 'short']
        ])
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_missing_file(self):
        self.assertFalse(self.generator.validate_csv("nonexistent.csv"))

if __name__ == '__main__':
    unittest.main()
