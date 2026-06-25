import unittest
import os
import csv
import tempfile
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_dir = tempfile.TemporaryDirectory()
        logging.basicConfig(level=logging.ERROR)

    def tearDown(self):
        self.test_dir.cleanup()

    def create_def_file(self, rows):
        path = os.path.join(self.test_dir.name, "test_def.csv")
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Test', 'Test', '', '', '', '', '', '', ''])
            for row in rows:
                writer.writerow(row)
        return path

    def test_valid_file(self):
        rows = [
            ['1', '3', '40001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['2', '3', '40002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_def_file(rows)
        self.assertTrue(self.generator.validate_csv(path))

    def test_duplicate_tag(self):
        rows = [
            ['1', '3', '40001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['2', '3', '40002', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_def_file(rows)
        # Duplicate tag is fatal
        self.assertFalse(self.generator.validate_csv(path))

    def test_out_of_range_address(self):
        rows = [
            ['1', '3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_def_file(rows)
        self.assertFalse(self.generator.validate_csv(path))

    def test_address_overlap(self):
        rows = [
            ['1', '3', '40001', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['2', '3', '40002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_def_file(rows)
        # Overlap is a warning in Generator, but let's see if validate_csv still returns True or False
        # Currently _check_address_overlap logs warning. validate_csv only sets is_valid = False for fatal errors like tag or invalid address format/range.
        # Wait, let's check validate_csv again.
        # It doesn't set is_valid = False for overlaps.
        self.assertTrue(self.generator.validate_csv(path))

    def test_invalid_address_format(self):
        rows = [
            ['1', '3', 'invalid', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_def_file(rows)
        self.assertFalse(self.generator.validate_csv(path))

if __name__ == '__main__':
    unittest.main()
