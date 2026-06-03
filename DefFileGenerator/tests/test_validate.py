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

    def tearDown(self):
        self.test_dir.cleanup()

    def create_csv(self, rows):
        path = os.path.join(self.test_dir.name, 'test.csv')
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                # index, info1, info2, info3, info4, name, tag, coefa, coefb, map, action
                writer.writerow([str(i)] + row)
        return path

    def test_validate_success(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'W', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'W', '4']
        ]
        path = self.create_csv(rows)
        self.assertTrue(self.generator.validate_csv(path))

    def test_validate_duplicate_tag(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'W', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'W', '4']
        ]
        path = self.create_csv(rows)
        with self.assertLogs(level='ERROR') as cm:
            self.assertFalse(self.generator.validate_csv(path))
            self.assertTrue(any("Duplicate Tag 'tag1'" in line for line in cm.output))

    def test_validate_address_overlap(self):
        rows = [
            ['3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'W', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'W', '4']
        ]
        path = self.create_csv(rows)
        with self.assertLogs(level='WARNING') as cm:
            # Overlap is currently a warning, so validate_csv should still return True if no fatal errors
            self.assertTrue(self.generator.validate_csv(path))
            self.assertTrue(any("Address overlap detected for 'Var2' at 101" in line for line in cm.output))

    def test_validate_hex_address(self):
        rows = [
            ['3', '0x100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'W', '4'],
            ['3', '256', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'W', '4']
        ]
        path = self.create_csv(rows)
        # 0x100 is 256, so this should trigger an overlap warning
        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_csv(path))
            self.assertTrue(any("Address overlap detected for 'Var2' at 256" in line for line in cm.output))

    def test_validate_malformed_row(self):
        path = os.path.join(self.test_dir.name, 'malformed.csv')
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['header'])
            writer.writerow(['1', '3', '100']) # Too few columns

        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_csv(path))
            self.assertTrue(any("Malformed row" in line for line in cm.output))

if __name__ == '__main__':
    unittest.main()
