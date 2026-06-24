import unittest
import os
import csv
import tempfile
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.temp_dir = tempfile.TemporaryDirectory()
        logging.basicConfig(level=logging.INFO)

    def tearDown(self):
        self.temp_dir.cleanup()

    def create_def_file(self, rows):
        path = os.path.join(self.temp_dir.name, "test_def.csv")
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)
        return path

    def test_valid_file(self):
        rows = [
            ['3', '40001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '40002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_def_file(rows)
        self.assertTrue(self.generator.validate_csv(path))

    def test_duplicate_tag(self):
        rows = [
            ['3', '40001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '40002', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_def_file(rows)
        self.assertFalse(self.generator.validate_csv(path))

    def test_address_overlap_warning(self):
        # Overlaps are logged as warnings, but validate_csv currently returns True for overlaps
        # unless it's a fatal error like duplicate tags.
        rows = [
            ['3', '40001', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '40002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_def_file(rows)
        # Should be valid (returns True) but log a warning
        self.assertTrue(self.generator.validate_csv(path))

    def test_invalid_address_range(self):
        rows = [
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_def_file(rows)
        self.assertFalse(self.generator.validate_csv(path))

    def test_compound_address_validation(self):
        rows = [
            ['3', '40001_20', 'STRING', '', 'Var1', 'tag1', '1.0', '0.0', '', '4'],
            ['3', '40030_0_1', 'BITS', '', 'Var2', 'tag2', '1.0', '0.0', '', '4']
        ]
        path = self.create_def_file(rows)
        self.assertTrue(self.generator.validate_csv(path))

if __name__ == '__main__':
    unittest.main()
