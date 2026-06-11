import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_csv = "test_validate.csv"
        # Setup logging to capture warnings/errors if needed, or just let them go to stderr
        logging.basicConfig(level=logging.INFO)

    def tearDown(self):
        if os.path.exists(self.test_csv):
            os.remove(self.test_csv)

    def _write_csv(self, rows):
        with open(self.test_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(["protocol", "category", "mfg", "model", "", "", "", "", "", "", ""])
            for row in rows:
                writer.writerow(row)

    def test_validate_valid(self):
        rows = [
            ["1", "3", "100", "U16", "", "Var1", "tag1", "1", "0", "V", "4"]
        ]
        self._write_csv(rows)
        self.assertTrue(self.generator.validate_csv(self.test_csv))

    def test_validate_duplicate_tag(self):
        rows = [
            ["1", "3", "100", "U16", "", "Var1", "tag1", "1", "0", "V", "4"],
            ["2", "3", "101", "U16", "", "Var2", "tag1", "1", "0", "V", "4"]
        ]
        self._write_csv(rows)
        self.assertFalse(self.generator.validate_csv(self.test_csv))

    def test_validate_address_overlap(self):
        rows = [
            ["1", "3", "100", "U32", "", "Var1", "tag1", "1", "0", "V", "4"],
            ["2", "3", "101", "U16", "", "Var2", "tag2", "1", "0", "V", "4"]
        ]
        self._write_csv(rows)
        # Overlap is a warning, so validate_csv should still return True unless I change it.
        # Wait, _check_address_overlap only logs warnings.
        self.assertTrue(self.generator.validate_csv(self.test_csv))

    def test_validate_invalid_address(self):
        rows = [
            ["1", "3", "INVALID", "U16", "", "Var1", "tag1", "1", "0", "V", "4"]
        ]
        self._write_csv(rows)
        self.assertFalse(self.generator.validate_csv(self.test_csv))

    def test_validate_address_out_of_range(self):
        rows = [
            ["1", "3", "70000", "U16", "", "Var1", "tag1", "1", "0", "V", "4"]
        ]
        self._write_csv(rows)
        # Range error is a warning, so it remains valid in terms of format.
        self.assertTrue(self.generator.validate_csv(self.test_csv))

    def test_validate_short_row(self):
        # Row with fewer than 11 columns
        with open(self.test_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(["header"])
            writer.writerow(["1", "3", "100"])
        self.assertTrue(self.generator.validate_csv(self.test_csv)) # Skips and remains valid

if __name__ == '__main__':
    unittest.main()
