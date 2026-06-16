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

    def write_test_csv(self, header, rows):
        with open(self.test_file, 'w', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(header)
            writer.writerows(rows)

    def test_validate_address_range(self):
        # Valid addresses
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        self.assertTrue(self.generator.validate_address("0x0", "U16"))
        self.assertTrue(self.generator.validate_address("0xFFFF", "U16"))

        # Invalid addresses (out of range)
        self.assertFalse(self.generator.validate_address("65536", "U16"))
        self.assertFalse(self.generator.validate_address("-1", "U16"))
        self.assertFalse(self.generator.validate_address("0x10000", "U16"))

    def test_validate_csv_basic(self):
        header = ["modbusRTU", "Inverter", "TestMFG", "TestModel", "", "", "", "", "", "", ""]
        rows = [
            ["1", "3", "100", "U16", "", "Var1", "tag1", "1.0", "0.0", "V", "4"],
            ["2", "3", "101", "U16", "", "Var2", "tag2", "1.0", "0.0", "A", "4"]
        ]
        self.write_test_csv(header, rows)
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_validate_csv_duplicate_tag(self):
        header = ["modbusRTU", "Inverter", "TestMFG", "TestModel", "", "", "", "", "", "", ""]
        rows = [
            ["1", "3", "100", "U16", "", "Var1", "tag1", "1.0", "0.0", "V", "4"],
            ["2", "3", "101", "U16", "", "Var2", "tag1", "1.0", "0.0", "A", "4"] # Duplicate tag
        ]
        self.write_test_csv(header, rows)
        # We expect it to return False because duplicate tags are fatal in our validation logic
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_validate_csv_overlap(self):
        header = ["modbusRTU", "Inverter", "TestMFG", "TestModel", "", "", "", "", "", "", ""]
        rows = [
            ["1", "3", "100", "U32", "", "Var1", "tag1", "1.0", "0.0", "V", "4"],
            ["2", "3", "101", "U16", "", "Var2", "tag2", "1.0", "0.0", "A", "4"] # Overlaps with second register of U32
        ]
        self.write_test_csv(header, rows)
        # Overlaps are warnings, so it should still return True but log warnings
        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("overlap detected" in output for output in cm.output))

if __name__ == '__main__':
    unittest.main()
