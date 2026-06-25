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

    def create_csv(self, rows, manufacturer="Test", model="Test"):
        filepath = os.path.join(self.test_dir.name, "test_def.csv")
        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', manufacturer, model, '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)
        return filepath

    def test_validate_csv_valid(self):
        rows = [
            ['3', '30001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '30002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        filepath = self.create_csv(rows)
        self.assertTrue(self.generator.validate_csv(filepath))

    def test_validate_csv_duplicate_tag(self):
        rows = [
            ['3', '30001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '30002', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4'] # Duplicate Tag
        ]
        filepath = self.create_csv(rows)
        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(self.generator.validate_csv(filepath))
            self.assertTrue(any("Duplicate Tag 'tag1'" in m for m in log.output))

    def test_validate_csv_address_range(self):
        # Invalid range
        rows = [['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']]
        filepath = self.create_csv(rows)
        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(self.generator.validate_csv(filepath))
            self.assertTrue(any("out of standard Modbus range" in m for m in log.output))

    def test_validate_address_range_edge_cases(self):
        self.assertTrue(self.generator.validate_address('0', 'U16'))
        self.assertTrue(self.generator.validate_address('65535', 'U16'))
        self.assertFalse(self.generator.validate_address('-1', 'U16'))
        self.assertFalse(self.generator.validate_address('65536', 'U16'))

        # Hex
        self.assertTrue(self.generator.validate_address('0xFFFF', 'U16'))
        self.assertFalse(self.generator.validate_address('0x10000', 'U16'))

    def test_validate_csv_overlap(self):
        rows = [
            ['3', '30001', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '30002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4'] # Overlaps with 30002 part of Var1
        ]
        filepath = self.create_csv(rows)
        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(self.generator.validate_csv(filepath)) # Overlap is warning, not fatal for validate_csv currently
            self.assertTrue(any("Address overlap detected" in m for m in log.output))

    def test_validate_csv_insufficient_columns(self):
        filepath = os.path.join(self.test_dir.name, "bad_cols.csv")
        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['header'])
            writer.writerow(['1', '3', '30001']) # Only 3 cols

        with self.assertLogs(level='WARNING') as log:
            # It will skip the row, but returns True if nothing else failed
            self.assertTrue(self.generator.validate_csv(filepath))
            self.assertTrue(any("Insufficient columns" in m for m in log.output))

if __name__ == '__main__':
    unittest.main()
