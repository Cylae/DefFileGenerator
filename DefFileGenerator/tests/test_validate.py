import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.test_file = "test_validation.csv"

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def create_csv(self, rows):
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Test', 'Test', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)

    def test_valid_csv(self):
        self.create_csv([
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ])
        self.assertTrue(Generator.validate_csv(self.test_file))

    def test_duplicate_tag(self):
        self.create_csv([
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        with self.assertLogs(level='ERROR') as cm:
            self.assertFalse(Generator.validate_csv(self.test_file))
            self.assertTrue(any("Duplicate Tag 'tag1'" in line for line in cm.output))

    def test_address_overlap(self):
        self.create_csv([
            ['3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ])
        with self.assertLogs(level='WARNING') as cm:
            # Overlap is currently a warning, so validation might still return True
            # unless we decided to make it fatal.
            Generator.validate_csv(self.test_file)
            self.assertTrue(any("Address overlap detected" in line for line in cm.output))

    def test_address_range_warning(self):
        self.create_csv([
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        with self.assertLogs(level='WARNING') as cm:
            Generator.validate_csv(self.test_file)
            self.assertTrue(any("outside the standard Modbus range" in line for line in cm.output))

if __name__ == '__main__':
    unittest.main()
