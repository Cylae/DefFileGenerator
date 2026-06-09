import unittest
import os
import csv
import tempfile
import logging
from DefFileGenerator.def_gen import Generator

class TestValidateCSV(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_dir = tempfile.TemporaryDirectory()
        self.test_file = os.path.join(self.test_dir.name, "test_def.csv")

    def tearDown(self):
        self.test_dir.cleanup()

    def create_def_file(self, rows, header=None):
        if header is None:
            header = ['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', '']
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(header)
            for i, row in enumerate(rows, start=1):
                # Row format: Index, Info1, Info2, Info3, Info4, Name, Tag, CoefA, CoefB, Unit, Action
                writer.writerow([str(i)] + list(row))

    def test_valid_file(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']
        ]
        self.create_def_file(rows)
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_duplicate_tag(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'A', '4']
        ]
        self.create_def_file(rows)
        with self.assertLogs(level='ERROR') as cm:
            self.assertFalse(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Duplicate Tag 'tag1'" in output for output in cm.output))

    def test_invalid_address(self):
        rows = [
            ['3', 'invalid', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.create_def_file(rows)
        with self.assertLogs(level='WARNING') as cm:
            self.assertFalse(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Invalid Address 'invalid'" in output for output in cm.output))

    def test_address_out_of_range(self):
        rows = [
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.create_def_file(rows)
        with self.assertLogs(level='WARNING') as cm:
            # Range check is a warning, doesn't invalidate the file currently
            self.assertTrue(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Address 70000 is out of standard Modbus range" in output for output in cm.output))

    def test_insufficient_columns(self):
        # We manually write a row with fewer columns
        header = ['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', '']
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(header)
            writer.writerow(['1', '3', '100', 'U16']) # Only 4 columns

        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Insufficient columns" in output for output in cm.output))

    def test_address_overlap(self):
        rows = [
            ['3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']
        ]
        self.create_def_file(rows)
        with self.assertLogs(level='WARNING') as cm:
            # Overlap is a warning, doesn't invalidate the file currently
            self.assertTrue(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Address overlap detected" in output for output in cm.output))

if __name__ == '__main__':
    unittest.main()
