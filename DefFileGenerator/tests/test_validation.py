import unittest
import os
import csv
import tempfile
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.temp_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.temp_dir.cleanup()

    def create_def_file(self, rows):
        path = os.path.join(self.temp_dir.name, 'test_def.csv')
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)
        return path

    def test_valid_file(self):
        rows = [
            ['3', '30001', 'U16', '', 'Var1', 'var1', '1.0', '0.0', 'V', '4'],
            ['3', '30002', 'I32', '', 'Var2', 'var2', '1.0', '0.0', 'A', '4']
        ]
        path = self.create_def_file(rows)
        self.assertTrue(self.generator.validate_csv(path))

    def test_duplicate_tag(self):
        rows = [
            ['3', '30001', 'U16', '', 'Var1', 'dup_tag', '1.0', '0.0', 'V', '4'],
            ['3', '30002', 'U16', '', 'Var2', 'dup_tag', '1.0', '0.0', 'A', '4']
        ]
        path = self.create_def_file(rows)
        self.assertFalse(self.generator.validate_csv(path))

    def test_address_out_of_range(self):
        self.assertFalse(self.generator.validate_address('65536', 'U16'))
        self.assertFalse(self.generator.validate_address('-1', 'U16'))
        self.assertTrue(self.generator.validate_address('0', 'U16'))
        self.assertTrue(self.generator.validate_address('65535', 'U16'))

    def test_invalid_address_format(self):
        self.assertFalse(self.generator.validate_address('invalid', 'U16'))
        self.assertFalse(self.generator.validate_address('30001_abc', 'STRING'))

    def test_str_synonym_validation(self):
        # STR20 should be treated like STRING
        self.assertTrue(self.generator.validate_address('30001_20', 'STR20'))
        self.assertFalse(self.generator.validate_address('30001', 'STR20'))

if __name__ == '__main__':
    unittest.main()
