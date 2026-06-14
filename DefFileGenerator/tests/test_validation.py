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

    def write_csv(self, rows, header=None):
        if header is None:
            header = ["modbusRTU", "Inverter", "Mfg", "Model", "", "", "", "", "", "", ""]
        with open(self.test_file, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(header)
            for row in rows:
                writer.writerow(row)

    def test_validate_valid(self):
        rows = [
            ["1", "3", "100", "U16", "", "Var1", "tag1", "1.0", "0.0", "V", "1"]
        ]
        self.write_csv(rows)
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_validate_duplicate_tag(self):
        rows = [
            ["1", "3", "100", "U16", "", "Var1", "tag1", "1.0", "0.0", "V", "1"],
            ["2", "3", "101", "U16", "", "Var2", "tag1", "1.0", "0.0", "V", "1"]
        ]
        self.write_csv(rows)
        # Disable logging for test
        logging.disable(logging.ERROR)
        self.assertFalse(self.generator.validate_csv(self.test_file))
        logging.disable(logging.NOTSET)

    def test_validate_address_range(self):
        # We want to check if it logs a warning but still returns True (or False if we decide range is fatal)
        # Current implementation: validate_address returns is_valid_format, range is just a warning.
        rows = [
            ["1", "3", "70000", "U16", "", "Var1", "tag1", "1.0", "0.0", "V", "1"]
        ]
        self.write_csv(rows)
        # Should be True because format is valid, even if range is out
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_validate_invalid_format(self):
        rows = [
            ["1", "3", "not_an_address", "U16", "", "Var1", "tag1", "1.0", "0.0", "V", "1"]
        ]
        self.write_csv(rows)
        logging.disable(logging.WARNING)
        self.assertFalse(self.generator.validate_csv(self.test_file))
        logging.disable(logging.NOTSET)

    def test_validate_short_row(self):
        self.write_csv([["1", "3"]])
        logging.disable(logging.WARNING)
        self.assertTrue(self.generator.validate_csv(self.test_file)) # Short rows are skipped, file might still be valid if empty or other rows are fine
        logging.disable(logging.NOTSET)

if __name__ == '__main__':
    unittest.main()
