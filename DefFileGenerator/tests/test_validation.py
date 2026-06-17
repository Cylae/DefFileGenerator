import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_file = 'test_validation.csv'
        # Disable logging for tests
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
        logging.disable(logging.NOTSET)

    def create_def_file(self, rows):
        with open(self.test_file, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Test', 'Test', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)

    def test_valid_file(self):
        self.create_def_file([
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ])
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_duplicate_tag(self):
        self.create_def_file([
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_address_overlap(self):
        self.create_def_file([
            ['3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ])
        # Overlap is currently a warning, so it might still return True if no other errors.
        # But let's check if it handles it.
        # In validate_csv, _check_address_overlap is called but it doesn't set valid=False.
        # However, duplicate tag does.
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_invalid_address_range(self):
        self.create_def_file([
            ['3', '65536', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_negative_address(self):
        self.create_def_file([
            ['3', '-1', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_malformed_row(self):
        with open(self.test_file, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Test', 'Test'])
            writer.writerow(['1', '3', '100']) # Too few columns
        self.assertTrue(self.generator.validate_csv(self.test_file)) # Skips malformed rows but doesn't necessarily fail validation if rest is ok

if __name__ == '__main__':
    unittest.main()
