import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.test_file = 'test_validate.csv'
        self.gen = Generator()

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def create_def_file(self, rows):
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)

    def test_valid_file(self):
        self.create_def_file([
            ['3', '40001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '40002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ])
        self.assertTrue(Generator.validate_csv(self.test_file))

    def test_duplicate_tag(self):
        self.create_def_file([
            ['3', '40001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '40002', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        with self.assertLogs(level='ERROR') as cm:
            self.assertFalse(Generator.validate_csv(self.test_file))
        self.assertTrue(any("Duplicate Tag 'tag1'" in line for line in cm.output))

    def test_address_overlap(self):
        self.create_def_file([
            ['3', '40001', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '40002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ])
        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(Generator.validate_csv(self.test_file)) # Overlap is warning
        self.assertTrue(any("Address overlap detected" in line for line in cm.output))

    def test_out_of_range_address(self):
        self.create_def_file([
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(Generator.validate_csv(self.test_file))
        self.assertTrue(any("outside standard Modbus range" in line for line in cm.output))

if __name__ == '__main__':
    unittest.main()
