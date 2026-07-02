import unittest
import csv
import os
import tempfile
from DefFileGenerator.def_gen import Generator

class TestValidateCSV(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.temp_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.temp_dir.cleanup()

    def create_csv(self, rows):
        path = os.path.join(self.temp_dir.name, 'test.csv')
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Test', 'Model', '', '', '', '', '', '', ''])
            for row in rows:
                writer.writerow(row)
        return path

    def test_valid_csv(self):
        rows = [
            ['1', '3', '40001', 'U16', '', 'Name1', 'tag1', '1.0', '0.0', 'W', '4']
        ]
        path = self.create_csv(rows)
        self.assertTrue(self.generator.validate_csv(path))

    def test_duplicate_tag(self):
        rows = [
            ['1', '3', '40001', 'U16', '', 'Name1', 'tag1', '1.0', '0.0', 'W', '4'],
            ['2', '3', '40002', 'U16', '', 'Name2', 'tag1', '1.0', '0.0', 'W', '4']
        ]
        path = self.create_csv(rows)
        self.assertFalse(self.generator.validate_csv(path))

    def test_address_overlap(self):
        rows = [
            ['1', '3', '40001', 'U32', '', 'Name1', 'tag1', '1.0', '0.0', 'W', '4'],
            ['2', '3', '40002', 'U16', '', 'Name2', 'tag2', '1.0', '0.0', 'W', '4']
        ]
        path = self.create_csv(rows)
        self.assertFalse(self.generator.validate_csv(path))

    def test_invalid_address(self):
        rows = [
            ['1', '3', 'invalid', 'U16', '', 'Name1', 'tag1', '1.0', '0.0', 'W', '4']
        ]
        path = self.create_csv(rows)
        self.assertFalse(self.generator.validate_csv(path))

    def test_insufficient_columns(self):
        path = os.path.join(self.temp_dir.name, 'short.csv')
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter'])
            writer.writerow(['1', '3', '40001']) # too short
        # Should log warning and return True if no other errors,
        # but since there are no valid rows, it might be trivial.
        self.assertTrue(self.generator.validate_csv(path))

if __name__ == '__main__':
    unittest.main()
