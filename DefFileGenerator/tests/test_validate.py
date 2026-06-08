import unittest
import os
import csv
import tempfile
import logging
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_dir = tempfile.TemporaryDirectory()
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        self.test_dir.cleanup()
        logging.disable(logging.NOTSET)

    def create_def_file(self, rows):
        path = os.path.join(self.test_dir.name, 'test_def.csv')
        with open(path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)
        return path

    def test_validate_success(self):
        # Index, Info1, Info2, Info3, Info4, Name, Tag, CoefA, CoefB, Unit, Action
        path = self.create_def_file([
            ['3', '30001', 'U16', '', 'Var1', 'var1', '1.0', '0.0', 'V', '4'],
            ['3', '30002', 'U16', '', 'Var2', 'var2', '1.0', '0.0', 'A', '4']
        ])
        self.assertTrue(self.generator.validate_csv(path))

    def test_validate_duplicate_tag(self):
        path = self.create_def_file([
            ['3', '30001', 'U16', '', 'Var1', 'dup_tag', '1.0', '0.0', 'V', '4'],
            ['3', '30002', 'U16', '', 'Var2', 'dup_tag', '1.0', '0.0', 'A', '4']
        ])
        self.assertFalse(self.generator.validate_csv(path))

    def test_validate_address_overlap(self):
        path = self.create_def_file([
            ['3', '30001', 'U32', '', 'Var1', 'var1', '1.0', '0.0', 'V', '4'],
            ['3', '30002', 'U16', '', 'Var2', 'var2', '1.0', '0.0', 'A', '4']
        ])
        # Overlap is a warning, not a fatal error
        self.assertTrue(self.generator.validate_csv(path))

    def test_validate_invalid_address(self):
        path = self.create_def_file([
            ['3', 'INVALID', 'U16', '', 'Var1', 'var1', '1.0', '0.0', 'V', '4']
        ])
        self.assertFalse(self.generator.validate_csv(path))

    def test_validate_out_of_range_address(self):
        path = self.create_def_file([
            ['3', '70000', 'U16', '', 'Var1', 'var1', '1.0', '0.0', 'V', '4']
        ])
        # Range is a warning, not a fatal error
        self.assertTrue(self.generator.validate_csv(path))

    def test_validate_nonexistent_file(self):
        self.assertFalse(self.generator.validate_csv('nonexistent.csv'))

    def test_validate_empty_file(self):
        path = os.path.join(self.test_dir.name, 'empty.csv')
        with open(path, 'w') as f:
            pass
        self.assertFalse(self.generator.validate_csv(path))

if __name__ == '__main__':
    unittest.main()
