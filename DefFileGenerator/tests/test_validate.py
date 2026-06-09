import unittest
import os
import csv
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.gen = Generator()
        self.test_file = "test_validate_def.csv"

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def create_def_file(self, rows):
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)

    def test_validate_success(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']
        ]
        self.create_def_file(rows)
        self.assertTrue(self.gen.validate_csv(self.test_file))

    def test_validate_duplicate_tag(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'A', '4']
        ]
        self.create_def_file(rows)
        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(self.gen.validate_csv(self.test_file))
            self.assertTrue(any("Duplicate Tag 'tag1'" in m for m in log.output))

    def test_validate_address_overlap(self):
        rows = [
            ['3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']
        ]
        self.create_def_file(rows)
        with self.assertLogs(level='WARNING') as log:
            # Overlap is currently a warning in Generator, not a fatal error for validate_csv
            self.assertTrue(self.gen.validate_csv(self.test_file))
            self.assertTrue(any("Address overlap detected" in m for m in log.output))

    def test_validate_invalid_address(self):
        rows = [
            ['3', 'invalid', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.create_def_file(rows)
        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(self.gen.validate_csv(self.test_file))
            self.assertTrue(any("Invalid Address 'invalid'" in m for m in log.output))

    def test_validate_out_of_range_address(self):
        rows = [
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.create_def_file(rows)
        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(self.gen.validate_csv(self.test_file))
            self.assertTrue(any("outside standard Modbus range" in m for m in log.output))

    def test_validate_short_row(self):
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel'])
            writer.writerow(['1', '3', '100']) # Too short
        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(self.gen.validate_csv(self.test_file))
            self.assertTrue(any("Row has fewer than 11 columns" in m for m in log.output))

if __name__ == "__main__":
    unittest.main()
