import unittest
import os
import logging
import csv
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_file = "test_validate.csv"

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def write_csv(self, rows):
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            for row in rows:
                writer.writerow(row)

    def test_validate_success(self):
        rows = [
            ['1', '3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['2', '3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        self.write_csv(rows)
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_duplicate_tag(self):
        rows = [
            ['1', '3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['2', '3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.write_csv(rows)
        with self.assertLogs(level='ERROR') as cm:
            self.assertFalse(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Duplicate Tag 'tag1'" in msg for msg in cm.output))

    def test_address_overlap(self):
        rows = [
            ['1', '3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['2', '3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        self.write_csv(rows)
        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Address overlap detected" in msg for msg in cm.output))

    def test_address_out_of_range(self):
        rows = [
            ['1', '3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.write_csv(rows)
        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("out of standard range" in msg for msg in cm.output))

if __name__ == '__main__':
    unittest.main()
