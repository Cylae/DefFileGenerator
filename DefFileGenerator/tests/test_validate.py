import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.test_csv = "test_validate.csv"

    def tearDown(self):
        if os.path.exists(self.test_csv):
            os.remove(self.test_csv)

    def _write_test_csv(self, rows):
        with open(self.test_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)

    def test_validate_csv_success(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        self._write_test_csv(rows)
        self.assertTrue(Generator.validate_csv(self.test_csv))

    def test_validate_csv_duplicate_tag(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'dup_tag', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'dup_tag', '1.0', '0.0', 'V', '4']
        ]
        self._write_test_csv(rows)
        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(Generator.validate_csv(self.test_csv))
            self.assertTrue(any("FATAL - Duplicate Tag detected" in m for m in log.output))

    def test_validate_csv_address_overlap(self):
        rows = [
            ['3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        self._write_test_csv(rows)
        with self.assertLogs(level='WARNING') as log:
            # Overlap is a warning, not fatal for validate_csv success unless we change requirements
            self.assertTrue(Generator.validate_csv(self.test_csv))
            self.assertTrue(any("Address overlap detected" in m for m in log.output))

    def test_validate_csv_invalid_range(self):
        rows = [
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self._write_test_csv(rows)
        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(Generator.validate_csv(self.test_csv))
            self.assertTrue(any("outside standard Modbus range" in m for m in log.output))

    def test_validate_csv_hex_address(self):
        rows = [
            ['3', '0x64', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'] # 0x64 = 100
        ]
        self._write_test_csv(rows)
        self.assertTrue(Generator.validate_csv(self.test_csv))

    def test_validate_csv_incomplete_row(self):
        with open(self.test_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel'])
            writer.writerow(['1', '3', '100']) # Incomplete
        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(Generator.validate_csv(self.test_csv))
            self.assertTrue(any("Incomplete row" in m for m in log.output))

if __name__ == '__main__':
    unittest.main()
