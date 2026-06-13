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

    def create_csv(self, rows):
        path = os.path.join(self.test_dir.name, "test_def.csv")
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)
        return path

    def test_valid_file(self):
        path = self.create_csv([
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '1'],
            ['4', '200', 'U32', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']
        ])
        self.assertTrue(self.generator.validate_csv(path))

    def test_duplicate_tag(self):
        path = self.create_csv([
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '1'],
            ['3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'A', '1']
        ])
        self.assertFalse(self.generator.validate_csv(path))

    def test_address_out_of_range(self):
        # We only log a warning for address range, but it should still be "valid"
        # unless it's an unparseable format.
        path = self.create_csv([
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '1']
        ])
        # It's currently implemented as a warning, so it returns True.
        self.assertTrue(self.generator.validate_csv(path))

    def test_invalid_address_format(self):
        path = self.create_csv([
            ['3', 'not_an_address', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '1']
        ])
        self.assertFalse(self.generator.validate_csv(path))

    def test_address_overlap(self):
        path = self.create_csv([
            ['3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '1'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '1']
        ])
        # Overlap is a warning, so it remains True.
        self.assertTrue(self.generator.validate_csv(path))

if __name__ == '__main__':
    unittest.main()
