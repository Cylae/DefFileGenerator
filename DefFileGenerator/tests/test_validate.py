import unittest
import os
import csv
import tempfile
import logging
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.temp_dir = tempfile.TemporaryDirectory()
        self.test_file = os.path.join(self.temp_dir.name, "test_def.csv")

    def tearDown(self):
        self.temp_dir.cleanup()

    def create_def_file(self, rows):
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)

    def test_valid_file(self):
        rows = [
            ['3', '40001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '40002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        self.create_def_file(rows)
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_duplicate_tag(self):
        rows = [
            ['3', '40001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '40002', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.create_def_file(rows)
        # Suppress error log for test
        with self.assertLogs(level='ERROR') as cm:
            self.assertFalse(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Duplicate Tag 'tag1' (Fatal)" in msg for msg in cm.output))

    def test_address_overlap(self):
        rows = [
            ['3', '40001', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '40002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        self.create_def_file(rows)
        with self.assertLogs(level='WARNING') as cm:
            # Overlap is a warning, so validate_csv should still return True (unless we decide otherwise)
            # Currently validate_csv returns is_valid which is only set to False on duplicate tags or invalid address format
            self.assertTrue(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Address overlap detected" in msg for msg in cm.output))

    def test_invalid_address_range(self):
        rows = [
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.create_def_file(rows)
        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Address 70000 is outside standard Modbus range" in msg for msg in cm.output))

    def test_short_row(self):
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '40001', 'U16']) # Too short

        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Row has fewer than 11 columns" in msg for msg in cm.output))

if __name__ == '__main__':
    unittest.main()
