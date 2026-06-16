import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_file = "test_validation.csv"
        logging.basicConfig(level=logging.ERROR)

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
            ['3', '30001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '1'],
            ['3', '30002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '1']
        ])
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_duplicate_tag(self):
        self.create_def_file([
            ['3', '30001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '1'],
            ['3', '30002', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'A', '1']
        ])
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_invalid_address(self):
        self.create_def_file([
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '1']
        ])
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_address_overlap(self):
        self.create_def_file([
            ['3', '30001', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '1'],
            ['3', '30002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '1']
        ])
        # Overlap detected, but validate_csv currently returns True for overlaps (logs warning)
        # unless it's a fatal error. Based on current implementation, it logs warning.
        # Let's verify it still passes if only warnings are issued.
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_insufficient_columns(self):
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel'])
            writer.writerow(['1', '3', '30001']) # too few
        self.assertTrue(self.generator.validate_csv(self.test_file)) # It logs warning and continues

if __name__ == '__main__':
    unittest.main()
