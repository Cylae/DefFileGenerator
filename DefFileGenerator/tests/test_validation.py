import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_file = "test_validate.csv"

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def write_test_csv(self, rows):
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)

    def test_validate_csv_valid(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        self.write_test_csv(rows)
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_validate_csv_duplicate_tag(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4'] # Duplicate tag1
        ]
        self.write_test_csv(rows)
        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("FATAL - Duplicate Tag 'tag1'" in m for m in log.output))

    def test_validate_csv_overlap(self):
        rows = [
            ['3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'], # 100, 101
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']  # 101 overlaps
        ]
        self.write_test_csv(rows)
        with self.assertLogs(level='WARNING') as log:
            # Overlap is a warning, so validate_csv still returns True but logs warning
            self.assertTrue(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Address overlap detected" in m for m in log.output))

    def test_validate_csv_address_range(self):
        rows = [
            ['3', '65536', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '-1', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        self.write_test_csv(rows)
        # We expect validation to fail because addresses are out of range
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_validate_csv_invalid_type(self):
        rows = [
            ['3', '100', 'INVALID_TYPE', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.write_test_csv(rows)
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_validate_csv_compound_address(self):
        rows = [
            ['3', '100_20', 'STRING', '', 'Var1', 'tag1', '1.0', '0.0', '', '4'],
            ['3', '200_0_1', 'BITS', '', 'Var2', 'tag2', '1.0', '0.0', '', '4']
        ]
        self.write_test_csv(rows)
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_validate_csv_bom_handling(self):
        # Write with UTF-8 BOM
        with open(self.test_file, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])

        self.assertTrue(self.generator.validate_csv(self.test_file))

if __name__ == '__main__':
    unittest.main()
