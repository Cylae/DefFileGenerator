import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_file = "test_validate.csv"
        # Disable logging for tests to keep output clean
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
        logging.disable(logging.NOTSET)

    def create_def_file(self, rows):
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)

    def test_valid_file(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'I32', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']
        ]
        self.create_def_file(rows)
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_duplicate_tag(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'A', '4']
        ]
        self.create_def_file(rows)
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_address_overlap(self):
        # I32 uses 2 registers. 100 and 101.
        rows = [
            ['3', '100', 'I32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']
        ]
        self.create_def_file(rows)
        # Overlaps are now treated as fatal validation errors
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_invalid_type(self):
        rows = [
            ['3', '100', 'INVALID', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.create_def_file(rows)
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_invalid_address_format(self):
        rows = [
            ['3', '!!!', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.create_def_file(rows)
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_out_of_range_address(self):
        rows = [
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.create_def_file(rows)
        # Out of range is a warning, so it should be valid
        self.assertTrue(self.generator.validate_csv(self.test_file))

if __name__ == '__main__':
    unittest.main()
