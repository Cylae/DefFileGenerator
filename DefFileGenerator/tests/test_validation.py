import unittest
import os
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_file = "test_validation.csv"
        # Disable logging for tests
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
        logging.disable(logging.NOTSET)

    def write_csv(self, rows):
        with open(self.test_file, 'w', encoding='utf-8') as f:
            f.write("modbusRTU;Inverter;MFG;MODEL;;;;;;;\n")
            for row in rows:
                f.write(";".join(row) + "\n")

    def test_duplicate_tag(self):
        self.write_csv([
            ["1", "3", "40001", "U16", "", "Name1", "tag1", "1", "0", "V", "4"],
            ["2", "3", "40002", "U16", "", "Name2", "tag1", "1", "0", "V", "4"]
        ])
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_address_overlap(self):
        # Overlap should only be a warning, validate_csv currently returns True for overlaps
        # but let's check the logic. Actually, my implementation currently returns True for overlaps
        # and only False for duplicate tags or invalid addresses.
        self.write_csv([
            ["1", "3", "40001", "U32", "", "Name1", "tag1", "1", "0", "V", "4"],
            ["2", "3", "40002", "U16", "", "Name2", "tag2", "1", "0", "V", "4"]
        ])
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_invalid_address(self):
        self.write_csv([
            ["1", "3", "invalid", "U16", "", "Name1", "tag1", "1", "0", "V", "4"]
        ])
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_out_of_range_address(self):
        # Warning only
        self.write_csv([
            ["1", "3", "70000", "U16", "", "Name1", "tag1", "1", "0", "V", "4"]
        ])
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_valid_file(self):
        self.write_csv([
            ["1", "3", "40001", "U16", "", "Name1", "tag1", "1", "0", "V", "4"],
            ["2", "3", "40002", "U16", "", "Name2", "tag2", "1", "0", "V", "4"]
        ])
        self.assertTrue(self.generator.validate_csv(self.test_file))

if __name__ == "__main__":
    unittest.main()
