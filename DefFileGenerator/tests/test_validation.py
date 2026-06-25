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
            header = ["modbusRTU", "Inverter", "MFG", "MODEL", "", "", "", "", "", "", ""]
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(header)
            writer.writerows(rows)

    def test_validate_success(self):
        rows = [
            ["1", "3", "40001", "U16", "", "Name1", "tag1", "1.0", "0.0", "V", "4"],
            ["2", "3", "40002", "U16", "", "Name2", "tag2", "1.0", "0.0", "V", "4"]
        ]
        self.write_csv(rows)
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_validate_duplicate_tag(self):
        rows = [
            ["1", "3", "40001", "U16", "", "Name1", "tag1", "1.0", "0.0", "V", "4"],
            ["2", "3", "40002", "U16", "", "Name2", "tag1", "1.0", "0.0", "V", "4"]
        ]
        self.write_csv(rows)
        with self.assertLogs(level='ERROR') as cm:
            self.assertFalse(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Duplicate Tag 'tag1' (Fatal)" in msg for msg in cm.output))

    def test_validate_invalid_address(self):
        rows = [
            ["1", "3", "70000", "U16", "", "Name1", "tag1", "1.0", "0.0", "V", "4"]
        ]
        self.write_csv(rows)
        # Should return False due to range check
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_validate_overlap(self):
        rows = [
            ["1", "3", "40001", "U32", "", "Name1", "tag1", "1.0", "0.0", "V", "4"],
            ["2", "3", "40002", "U16", "", "Name2", "tag2", "1.0", "0.0", "V", "4"]
        ]
        self.write_csv(rows)
        with self.assertLogs(level='WARNING') as cm:
            # Overlap is a warning, so validate_csv should still return True if no fatal errors
            self.assertTrue(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Address overlap detected" in msg for msg in cm.output))

    def test_address_range_validation(self):
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        self.assertTrue(self.generator.validate_address("0x0", "U16"))
        self.assertTrue(self.generator.validate_address("0xFFFF", "U16"))

        with self.assertLogs(level='WARNING') as cm:
            self.assertFalse(self.generator.validate_address("65536", "U16"))
            self.assertFalse(self.generator.validate_address("-1", "U16"))

    def test_str_n_validation(self):
        self.assertTrue(self.generator.validate_address("30001_20", "STR20"))
        self.assertTrue(self.generator.validate_address("0x100_10", "STR10"))
        self.assertFalse(self.generator.validate_address("30001", "STR20")) # STR needs _len

if __name__ == '__main__':
    unittest.main()
