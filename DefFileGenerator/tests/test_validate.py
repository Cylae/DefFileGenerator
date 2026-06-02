import unittest
import os
import csv
import tempfile
import logging
from DefFileGenerator.def_gen import Generator
from DefFileGenerator.main import main
from unittest.mock import patch
import sys

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.test_dir.cleanup()

    def create_csv(self, rows, filename="test.csv"):
        path = os.path.join(self.test_dir.name, filename)
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            for row in rows:
                writer.writerow(row)
        return path

    def test_validate_csv_success(self):
        # Index, Info1, Info2, Info3, Info4, Name, Tag, CoefA, CoefB, Map, Action
        rows = [
            ['1', '3', '30001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['2', '3', '30002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']
        ]
        path = self.create_csv(rows)
        self.assertTrue(self.generator.validate_csv(path))

    def test_validate_csv_duplicate_tag(self):
        rows = [
            ['1', '3', '30001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['2', '3', '30002', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'A', '4']
        ]
        path = self.create_csv(rows)
        with self.assertLogs(level='ERROR') as cm:
            self.assertFalse(self.generator.validate_csv(path))
            self.assertTrue(any("Duplicate Tag 'tag1'" in msg for msg in cm.output))

    def test_validate_csv_address_overlap(self):
        rows = [
            ['1', '3', '30001', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['2', '3', '30002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']
        ]
        path = self.create_csv(rows)
        with self.assertLogs(level='WARNING') as cm:
            # U32 at 30001 uses 30001 and 30002.
            # U16 at 30002 overlaps.
            self.assertTrue(self.generator.validate_csv(path))
            self.assertTrue(any("Address overlap detected" in msg for msg in cm.output))

    def test_validate_csv_invalid_address_range(self):
        rows = [
            ['1', '3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_csv(rows)
        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_csv(path))
            self.assertTrue(any("Address 70000 is outside standard Modbus range" in msg for msg in cm.output))

    def test_validate_command_cli_success(self):
        rows = [['1', '3', '30001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']]
        path = self.create_csv(rows)
        test_args = ['main.py', 'validate', path]
        with patch.object(sys, 'argv', test_args):
            with self.assertLogs(level='INFO') as cm:
                main()
                self.assertTrue(any(f"Validation successful for {path}" in msg for msg in cm.output))

    def test_validate_command_cli_failure(self):
        # Duplicate tag causes failure
        rows = [
            ['1', '3', '30001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['2', '3', '30002', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'A', '4']
        ]
        path = self.create_csv(rows)
        test_args = ['main.py', 'validate', path]
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)

if __name__ == '__main__':
    unittest.main()
