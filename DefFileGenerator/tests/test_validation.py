import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_file = "test_validate.csv"
        # Disable logging for tests to keep output clean
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
        logging.disable(logging.NOTSET)

    def create_csv(self, rows):
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Test', 'Model', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + list(row))

    def test_valid_file(self):
        self.create_csv([
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ])
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_duplicate_tag(self):
        self.create_csv([
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_address_overlap(self):
        self.create_csv([
            ['3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ])
        # Overlap is currently a warning, so validate_csv should still return True
        # but let's check if it behaves as expected.
        # Wait, in validate_csv, _check_address_overlap is called but it doesn't set valid=False.
        # So it should be True.
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_invalid_address_range(self):
        self.create_csv([
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_invalid_address_format(self):
        self.create_csv([
            ['3', 'not_an_address', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_missing_tag(self):
        self.create_csv([
            ['3', '100', 'U16', '', 'Var1', '', '1.0', '0.0', 'V', '4']
        ])
        self.assertFalse(self.generator.validate_csv(self.test_file))

if __name__ == '__main__':
    unittest.main()
