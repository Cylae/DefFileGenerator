import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_file = 'test_validation.csv'

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def create_def_file(self, rows, header=None):
        if header is None:
            header = ['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', '']
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(header)
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)

    def test_validate_csv_valid(self):
        rows = [
            ['3', '30001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.create_def_file(rows)
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_validate_csv_duplicate_tag(self):
        rows = [
            ['3', '30001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '30002', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.create_def_file(rows)
        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Duplicate Tag 'tag1'" in m for m in log.output))

    def test_validate_csv_address_overlap(self):
        rows = [
            ['3', '30001', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '30002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        self.create_def_file(rows)
        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(self.generator.validate_csv(self.test_file)) # Overlap is warning, not fatal for validate_csv return
            self.assertTrue(any("Address overlap detected" in m for m in log.output))

    def test_validate_address_range(self):
        # 65535 is OK
        self.assertTrue(self.generator.validate_address('65535', 'U16'))

        # 65536 is outside range (logs warning)
        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(self.generator.validate_address('65536', 'U16'))
            self.assertTrue(any("Address 65536 is outside standard Modbus range" in m for m in log.output))

        # Negative is outside range
        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(self.generator.validate_address('-1', 'U16'))
            self.assertTrue(any("Address -1 is outside standard Modbus range" in m for m in log.output))

    def test_insufficient_columns(self):
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel'])
            writer.writerow(['1', '3', '30001', 'U16']) # Only 4 columns

        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Insufficient columns" in m for m in log.output))

if __name__ == '__main__':
    unittest.main()
