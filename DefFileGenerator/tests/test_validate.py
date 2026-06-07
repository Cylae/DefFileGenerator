import unittest
import os
import csv
import tempfile
import logging
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.test_file = os.path.join(self.temp_dir.name, "test_def.csv")
        # Silence logging during tests
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        self.temp_dir.cleanup()
        logging.disable(logging.NOTSET)

    def create_def_file(self, rows):
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            for row in rows:
                writer.writerow(row)

    def test_valid_file(self):
        self.create_def_file([
            ['1', '3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['2', '3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']
        ])
        self.assertTrue(Generator.validate_csv(self.test_file))

    def test_duplicate_tag(self):
        # Duplicate tags should be FATAL (False)
        self.create_def_file([
            ['1', '3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['2', '3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'A', '4']
        ])
        self.assertFalse(Generator.validate_csv(self.test_file))

    def test_address_overlap(self):
        # Overlaps are warnings, but validate_csv currently returns True for them
        # (unless we change it to return False for overlaps too)
        # For now, it detects them. Let's verify it doesn't crash.
        self.create_def_file([
            ['1', '3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['2', '3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']
        ])
        self.assertTrue(Generator.validate_csv(self.test_file))

    def test_invalid_address_range(self):
        self.create_def_file([
            ['1', '3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        # validate_address logs a warning, returns True if format is ok but range is off?
        # Actually my implementation of validate_address returns True even if it warns about range.
        self.assertTrue(Generator.validate_csv(self.test_file))

    def test_nonexistent_file(self):
        self.assertFalse(Generator.validate_csv("nonexistent_file.csv"))

    def test_empty_file(self):
        with open(self.test_file, 'w') as f:
            pass
        self.assertFalse(Generator.validate_csv(self.test_file))

if __name__ == '__main__':
    unittest.main()
