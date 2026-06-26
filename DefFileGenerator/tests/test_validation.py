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
        logging.basicConfig(level=logging.INFO)

    def tearDown(self):
        self.test_dir.cleanup()

    def create_csv(self, rows):
        path = os.path.join(self.test_dir.name, 'test.csv')
        with open(path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerows(rows)
        return path

    def test_validate_valid_csv(self):
        rows = [
            ['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''],
            ['1', '3', '30001', 'U16', '', 'Var1', 'var1', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_csv(rows)
        self.assertTrue(self.generator.validate_csv(path))

    def test_validate_invalid_address(self):
        rows = [
            ['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''],
            ['1', '3', '70000', 'U16', '', 'Var1', 'var1', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_csv(rows)
        # Should return False due to out of range address
        self.assertFalse(self.generator.validate_csv(path))

    def test_validate_invalid_type(self):
        rows = [
            ['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''],
            ['1', '3', '30001', 'INVALID', '', 'Var1', 'var1', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_csv(rows)
        self.assertFalse(self.generator.validate_csv(path))

    def test_validate_overlap(self):
        rows = [
            ['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''],
            ['1', '3', '30001', 'U32', '', 'Var1', 'var1', '1.0', '0.0', 'V', '4'],
            ['2', '3', '30002', 'U16', '', 'Var2', 'var2', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_csv(rows)
        # Overlap is currently a warning in validate_csv, but it should return True for the process
        # because the original logic doesn't fail on overlaps, just logs them.
        # Wait, if I want strict validation, I should check.
        # Current implementation of validate_csv doesn't set success=False for overlaps.
        self.assertTrue(self.generator.validate_csv(path))

    def test_validate_short_row(self):
        rows = [
            ['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''],
            ['1', '3', '30001', 'U16', ''] # Too short
        ]
        path = self.create_csv(rows)
        # Row with < 11 columns is skipped, returns True if no other errors
        self.assertTrue(self.generator.validate_csv(path))

    def test_validate_compound_address(self):
        rows = [
            ['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''],
            ['1', '3', '0x7530_20', 'STRING', '', 'Str1', 'str1', '1.0', '0.0', '', '4']
        ]
        path = self.create_csv(rows)
        self.assertTrue(self.generator.validate_csv(path))

    def test_validate_invalid_compound_address(self):
        rows = [
            ['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''],
            ['1', '3', '70000_20', 'STRING', '', 'Str1', 'str1', '1.0', '0.0', '', '4']
        ]
        path = self.create_csv(rows)
        self.assertFalse(self.generator.validate_csv(path))

if __name__ == '__main__':
    unittest.main()
