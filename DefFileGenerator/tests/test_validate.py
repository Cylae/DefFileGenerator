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
        logging.basicConfig(level=logging.INFO)

    def tearDown(self):
        self.test_dir.cleanup()

    def create_def_file(self, rows, manufacturer="TestMfg", model="TestModel"):
        path = os.path.join(self.test_dir.name, "test_def.csv")
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', manufacturer, model, '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + list(row))
        return path

    def test_valid_definition(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']
        ]
        path = self.create_def_file(rows)
        self.assertTrue(self.generator.validate_csv(path))

    def test_duplicate_tag(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'dup_tag', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'dup_tag', '1.0', '0.0', 'A', '4']
        ]
        path = self.create_def_file(rows)
        # Should be invalid due to duplicate tag
        self.assertFalse(self.generator.validate_csv(path))

    def test_address_overlap(self):
        rows = [
            ['3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']
        ]
        path = self.create_def_file(rows)
        # Address 101 is used by both Var1 (100, 101) and Var2 (101)
        # validate_csv returns True but logs warning for overlaps (non-fatal in current impl, but we check if it runs)
        self.assertTrue(self.generator.validate_csv(path))

    def test_invalid_address_format(self):
        rows = [
            ['3', 'not_an_address', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_def_file(rows)
        self.assertFalse(self.generator.validate_csv(path))

    def test_address_out_of_range(self):
        rows = [
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_def_file(rows)
        # Should be valid but log warning
        self.assertTrue(self.generator.validate_csv(path))

    def test_insufficient_columns(self):
        path = os.path.join(self.test_dir.name, "short_cols.csv")
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL'])
            writer.writerow(['1', '3', '100', 'U16']) # Only 4 columns
        self.assertTrue(self.generator.validate_csv(path)) # Skips row, file still "valid" as in no fatal errors found

if __name__ == '__main__':
    unittest.main()
