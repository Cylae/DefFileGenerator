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
        with open(path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)
        return path

    def test_valid_file(self):
        path = self.create_csv([
            ['3', '40001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '40002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ])
        self.assertTrue(self.generator.validate_csv(path))

    def test_duplicate_tag(self):
        path = self.create_csv([
            ['3', '40001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '40002', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        self.assertFalse(self.generator.validate_csv(path))

    def test_invalid_address(self):
        path = self.create_csv([
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        self.assertFalse(self.generator.validate_csv(path))

    def test_invalid_type(self):
        path = self.create_csv([
            ['3', '40001', 'INVALID', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        self.assertFalse(self.generator.validate_csv(path))

if __name__ == '__main__':
    unittest.main()
