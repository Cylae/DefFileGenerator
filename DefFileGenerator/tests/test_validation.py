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

    def tearDown(self):
        self.test_dir.cleanup()

    def test_validate_csv_valid(self):
        path = os.path.join(self.test_dir.name, "valid.csv")
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(["modbusRTU", "Inverter", "Mfg", "Model", "", "", "", "", "", "", ""])
            writer.writerow(["1", "3", "30001", "U16", "", "Name1", "tag1", "1.0", "0.0", "V", "4"])

        self.assertTrue(self.generator.validate_csv(path))

    def test_validate_csv_duplicate_tag(self):
        path = os.path.join(self.test_dir.name, "dup_tag.csv")
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(["modbusRTU", "Inverter", "Mfg", "Model", "", "", "", "", "", "", ""])
            writer.writerow(["1", "3", "30001", "U16", "", "Name1", "tag1", "1.0", "0.0", "V", "4"])
            writer.writerow(["2", "3", "30002", "U16", "", "Name2", "tag1", "1.0", "0.0", "V", "4"])

        # Suppress error logging for test
        logging.getLogger().setLevel(logging.CRITICAL)
        self.assertFalse(self.generator.validate_csv(path))
        logging.getLogger().setLevel(logging.INFO)

    def test_validate_address_range(self):
        # Range check is a warning, not a fatal error for validate_address itself currently,
        # but let's check it.
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_address("65536", "U16"))
            self.assertIn("outside standard Modbus range", cm.output[0])

if __name__ == '__main__':
    unittest.main()
