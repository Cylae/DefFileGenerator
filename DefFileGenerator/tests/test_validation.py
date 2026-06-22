import unittest
import os
import csv
import logging
import tempfile
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.temp_dir = tempfile.TemporaryDirectory()
        logging.basicConfig(level=logging.ERROR) # Suppress warnings during tests

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_validate_address_range(self):
        # Valid addresses
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        self.assertTrue(self.generator.validate_address("0x0", "U16"))
        self.assertTrue(self.generator.validate_address("0xFFFF", "U16"))

        # Out of range
        self.assertFalse(self.generator.validate_address("65536", "U16"))
        self.assertFalse(self.generator.validate_address("-1", "U16"))
        self.assertFalse(self.generator.validate_address("0x10000", "U16"))

    def test_validate_address_str_synonym(self):
        # STR<n> should be handled by validate_address
        self.assertTrue(self.generator.validate_address("100_20", "STR20"))
        self.assertTrue(self.generator.validate_address("100_20", "STRING"))
        self.assertFalse(self.generator.validate_address("100", "STR20")) # STR needs _len

    def test_validate_csv_basic(self):
        path = os.path.join(self.temp_dir.name, "test_def.csv")
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(["modbusRTU", "Inverter", "Mfg", "Model", "", "", "", "", "", "", ""])
            writer.writerow(["1", "3", "100", "U16", "", "Var1", "tag1", "1.0", "0.0", "V", "4"])

        self.assertTrue(self.generator.validate_csv(path))

    def test_validate_csv_duplicate_tag(self):
        path = os.path.join(self.temp_dir.name, "dup_tag.csv")
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(["modbusRTU", "Inverter", "Mfg", "Model", "", "", "", "", "", "", ""])
            writer.writerow(["1", "3", "100", "U16", "", "Var1", "tag1", "1.0", "0.0", "V", "4"])
            writer.writerow(["2", "3", "101", "U16", "", "Var2", "tag1", "1.0", "0.0", "V", "4"])

        self.assertFalse(self.generator.validate_csv(path))

    def test_validate_csv_address_overlap(self):
        path = os.path.join(self.temp_dir.name, "overlap.csv")
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(["modbusRTU", "Inverter", "Mfg", "Model", "", "", "", "", "", "", ""])
            # U32 takes 2 registers: 100, 101
            writer.writerow(["1", "3", "100", "U32", "", "Var1", "tag1", "1.0", "0.0", "V", "4"])
            writer.writerow(["2", "3", "101", "U16", "", "Var2", "tag2", "1.0", "0.0", "V", "4"])

        # Overlap is currently a warning, but let's see if validate_csv still returns True
        # (It should, unless we want overlaps to be fatal. Plan says detection, doesn't specify fatal).
        # Wait, the code for validate_csv does NOT set valid=False for overlaps.
        self.assertTrue(self.generator.validate_csv(path))

    def test_validate_csv_invalid_address(self):
        path = os.path.join(self.temp_dir.name, "invalid_addr.csv")
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(["modbusRTU", "Inverter", "Mfg", "Model", "", "", "", "", "", "", ""])
            writer.writerow(["1", "3", "70000", "U16", "", "Var1", "tag1", "1.0", "0.0", "V", "4"])

        self.assertFalse(self.generator.validate_csv(path))

if __name__ == "__main__":
    unittest.main()
