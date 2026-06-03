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
        self.test_file = os.path.join(self.test_dir.name, "test_def.csv")

    def tearDown(self):
        self.test_dir.cleanup()

    def create_def_file(self, rows):
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)

    def test_validate_csv_success(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        self.create_def_file(rows)
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_validate_csv_duplicate_tag(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.create_def_file(rows)
        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Duplicate Tag 'tag1' detected" in m for m in log.output))

    def test_validate_csv_address_overlap(self):
        rows = [
            ['3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'], # 100, 101
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']  # 101 (Overlap)
        ]
        self.create_def_file(rows)
        with self.assertLogs(level='WARNING') as log:
            # validate_csv returns True on overlaps (they are warnings), but logs them
            self.assertTrue(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Address overlap detected" in m for m in log.output))

    def test_validate_csv_out_of_range_address(self):
        rows = [
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.create_def_file(rows)
        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("outside standard Modbus range" in m for m in log.output))

if __name__ == '__main__':
    unittest.main()
