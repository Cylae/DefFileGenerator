#!/usr/bin/env python3
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
        logging.basicConfig(level=logging.ERROR)

    def tearDown(self):
        self.test_dir.cleanup()

    def create_csv(self, rows):
        path = os.path.join(self.test_dir.name, 'test_def.csv')
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Test', 'Model', '', '', '', '', '', '', ''])
            for idx, row in enumerate(rows, start=1):
                writer.writerow([str(idx)] + list(row))
        return path

    def test_valid_definition(self):
        path = self.create_csv([
            ['3', '30001', 'U16', '', 'Var1', 'var1', '1.0', '0.0', 'V', '4'],
            ['3', '30002', 'U16', '', 'Var2', 'var2', '1.0', '0.0', 'V', '4']
        ])
        self.assertTrue(self.generator.validate_csv(path))

    def test_duplicate_tag(self):
        path = self.create_csv([
            ['3', '30001', 'U16', '', 'Var1', 'dup_tag', '1.0', '0.0', 'V', '4'],
            ['3', '30002', 'U16', '', 'Var2', 'dup_tag', '1.0', '0.0', 'V', '4']
        ])
        self.assertFalse(self.generator.validate_csv(path))

    def test_address_out_of_range(self):
        path = self.create_csv([
            ['3', '65536', 'U16', '', 'Var1', 'var1', '1.0', '0.0', 'V', '4']
        ])
        self.assertFalse(self.generator.validate_csv(path))

        path = self.create_csv([
            ['3', '-1', 'U16', '', 'Var1', 'var1', '1.0', '0.0', 'V', '4']
        ])
        self.assertFalse(self.generator.validate_csv(path))

    def test_address_overlap(self):
        # 32-bit register at 30001 uses 30001 and 30002
        path = self.create_csv([
            ['3', '30001', 'U32', '', 'Var1', 'var1', '1.0', '0.0', 'V', '4'],
            ['3', '30002', 'U16', '', 'Var2', 'var2', '1.0', '0.0', 'V', '4']
        ])
        # Overlap is currently a warning in _check_address_overlap but validate_csv should still return True if it's just warnings
        # Wait, looking at validate_csv implementation: it doesn't set valid=False on overlaps.
        self.assertTrue(self.generator.validate_csv(path))

    def test_invalid_address_format(self):
        path = self.create_csv([
            ['3', 'invalid', 'U16', '', 'Var1', 'var1', '1.0', '0.0', 'V', '4']
        ])
        self.assertFalse(self.generator.validate_csv(path))

if __name__ == '__main__':
    unittest.main()
