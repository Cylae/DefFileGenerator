import unittest
import os
import logging
import csv
import tempfile
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.temp_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.temp_dir.cleanup()

    def create_csv(self, rows):
        path = os.path.join(self.temp_dir.name, 'test_def.csv')
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                # index, info1, info2, info3, info4, name, tag, coefa, coefb, unit, action
                full_row = [str(i)] + list(row)
                writer.writerow(full_row)
        return path

    def test_valid_file(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_csv(rows)
        self.assertTrue(self.generator.validate_csv(path))

    def test_duplicate_tag(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_csv(rows)
        with self.assertLogs(level='ERROR') as cm:
            result = self.generator.validate_csv(path)
            self.assertFalse(result)
            self.assertTrue(any("Duplicate Tag 'tag1'" in msg for msg in cm.output))

    def test_address_overlap(self):
        rows = [
            ['3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_csv(rows)
        with self.assertLogs(level='WARNING') as cm:
            self.generator.validate_csv(path)
            self.assertTrue(any("Address overlap detected" in msg for msg in cm.output))

    def test_address_out_of_range(self):
        rows = [
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_csv(rows)
        with self.assertLogs(level='WARNING') as cm:
            self.generator.validate_csv(path)
            self.assertTrue(any("Address 70000 out of range" in msg for msg in cm.output))

if __name__ == '__main__':
    unittest.main()
