import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_csv = 'test_validation.csv'

    def tearDown(self):
        if os.path.exists(self.test_csv):
            os.remove(self.test_csv)

    def create_csv(self, rows):
        with open(self.test_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)

    def test_validate_address_range(self):
        # Valid range
        self.assertTrue(self.generator.validate_address('0', 'U16'))
        self.assertTrue(self.generator.validate_address('65535', 'U16'))
        self.assertTrue(self.generator.validate_address('0x0', 'U16'))
        self.assertTrue(self.generator.validate_address('0xFFFF', 'U16'))

        # Out of range
        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(self.generator.validate_address('65536', 'U16'))
            self.assertTrue(any("out of standard Modbus range" in m for m in log.output))

        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(self.generator.validate_address('-1', 'U16'))
            self.assertTrue(any("out of standard Modbus range" in m for m in log.output))

    def test_validate_csv_success(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        self.create_csv(rows)
        self.assertTrue(self.generator.validate_csv(self.test_csv))

    def test_validate_csv_duplicate_tag(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.create_csv(rows)
        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(self.generator.validate_csv(self.test_csv))
            self.assertTrue(any("Duplicate Tag 'tag1'" in m for m in log.output))

    def test_validate_csv_address_overlap(self):
        rows = [
            ['3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'], # 100, 101
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']  # 101 (Overlap)
        ]
        self.create_csv(rows)
        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(self.generator.validate_csv(self.test_csv))
            self.assertTrue(any("Address overlap detected" in m for m in log.output))

    def test_validate_csv_invalid_address(self):
        rows = [
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.create_csv(rows)
        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(self.generator.validate_csv(self.test_csv))
            self.assertTrue(any("Invalid address format/range '70000'" in m for m in log.output))

    def test_validate_csv_short_row(self):
        # Row with only 5 columns (index + 4)
        with open(self.test_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['header'])
            writer.writerow(['1', '3', '100', 'U16', ''])

        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(self.generator.validate_csv(self.test_csv))
            self.assertTrue(any("Row has fewer than 11 columns" in m for m in log.output))

if __name__ == '__main__':
    unittest.main()
