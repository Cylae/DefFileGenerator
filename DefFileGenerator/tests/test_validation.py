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

    def tearDown(self):
        self.test_dir.cleanup()

    def create_def_file(self, rows, header=None):
        path = os.path.join(self.test_dir.name, 'test_def.csv')
        if header is None:
            header = ['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', '']
        with open(path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(header)
            for i, row in enumerate(rows, start=1):
                # Ensure row has 11 columns
                full_row = [str(i)] + list(row)
                while len(full_row) < 11:
                    full_row.append('')
                writer.writerow(full_row)
        return path

    def test_validate_csv_valid(self):
        rows = [
            ['3', '30001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '30002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_def_file(rows)
        self.assertTrue(self.generator.validate_csv(path))

    def test_validate_csv_duplicate_tag(self):
        rows = [
            ['3', '30001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '30002', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4'] # Duplicate tag
        ]
        path = self.create_def_file(rows)
        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(self.generator.validate_csv(path))
            self.assertTrue(any("Duplicate Tag 'tag1'" in m for m in log.output))

    def test_validate_csv_address_overlap(self):
        rows = [
            ['3', '30001', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'], # 30001, 30002
            ['3', '30002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']  # Overlap at 30002
        ]
        path = self.create_def_file(rows)
        # Overlaps are logged as warnings in _check_address_overlap
        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(self.generator.validate_csv(path))
            self.assertTrue(any("Address overlap detected" in m for m in log.output))

    def test_validate_csv_out_of_range_address(self):
        rows = [
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_def_file(rows)
        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(self.generator.validate_csv(path))
            self.assertTrue(any("out of standard Modbus range" in m for m in log.output))

    def test_validate_csv_invalid_header(self):
        path = self.create_def_file([], header=['only', 'three', 'cols'])
        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(self.generator.validate_csv(path))
            self.assertTrue(any("Invalid Webdyn definition header" in m for m in log.output))

    def test_validate_csv_short_row(self):
        path = os.path.join(self.test_dir.name, 'short_row.csv')
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'M', 'M', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '30001']) # Too short

        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(self.generator.validate_csv(path))
            self.assertTrue(any("Row has too few columns" in m for m in log.output))

if __name__ == '__main__':
    unittest.main()
