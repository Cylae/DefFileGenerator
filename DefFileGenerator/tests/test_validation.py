import unittest
import os
import csv
import logging
import tempfile
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        # Suppress logging during tests
        logging.disable(logging.CRITICAL)
        self.test_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        logging.disable(logging.NOTSET)
        self.test_dir.cleanup()

    def create_csv(self, rows, header=None):
        path = os.path.join(self.test_dir.name, "test_val.csv")
        if header is None:
            header = ["modbusRTU", "Inverter", "TestMFG", "TestModel", "", "", "", "", "", "", ""]
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(header)
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + list(row))
        return path

    def test_validate_address_range(self):
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        self.assertFalse(self.generator.validate_address("65536", "U16"))
        self.assertFalse(self.generator.validate_address("-1", "U16"))
        # Hex
        self.assertTrue(self.generator.validate_address("0xFFFF", "U16"))
        self.assertFalse(self.generator.validate_address("0x10000", "U16"))

    def test_validate_csv_success(self):
        rows = [
            ["3", "30001", "U16", "", "Voltage", "v_tag", "1.0", "0.0", "V", "4"],
            ["3", "30002", "U16", "", "Current", "c_tag", "1.0", "0.0", "A", "4"]
        ]
        path = self.create_csv(rows)
        self.assertTrue(self.generator.validate_csv(path))

    def test_validate_csv_duplicate_tag(self):
        rows = [
            ["3", "30001", "U16", "", "Voltage", "dup_tag", "1.0", "0.0", "V", "4"],
            ["3", "30002", "U16", "", "Current", "dup_tag", "1.0", "0.0", "A", "4"]
        ]
        path = self.create_csv(rows)
        # validate_csv should return False for duplicate tags
        self.assertFalse(self.generator.validate_csv(path))

    def test_validate_csv_address_overlap(self):
        rows = [
            ["3", "30001", "U32", "", "Power", "p_tag", "1.0", "0.0", "W", "4"],
            ["3", "30002", "U16", "", "Freq", "f_tag", "1.0", "0.0", "Hz", "4"]
        ]
        path = self.create_csv(rows)
        # Overlap (30001 is 2 regs: 30001, 30002) is fatal by default in my new implementation
        self.assertFalse(self.generator.validate_csv(path))

    def test_validate_csv_invalid_address(self):
        rows = [
            ["3", "70000", "U16", "", "Invalid", "inv_tag", "1.0", "0.0", "V", "4"]
        ]
        path = self.create_csv(rows)
        self.assertFalse(self.generator.validate_csv(path))

    def test_validate_csv_compound_address(self):
        # String address validation
        self.assertTrue(self.generator.validate_address("30030_20", "STRING"))
        self.assertTrue(self.generator.validate_address("30030_20", "STR20"))

        rows = [
            ["3", "30030_20", "STR20", "", "Serial", "s_tag", "1.0", "0.0", "", "4"]
        ]
        path = self.create_csv(rows)
        self.assertTrue(self.generator.validate_csv(path))

    def test_intelligent_action_defaulting(self):
        # Input Register (4) should default to 4 (Read Only)
        rows = [{'Name': 'InputVar', 'Address': '100', 'Type': 'U16', 'RegisterType': 'Input Register', 'Action': ''}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4')

        # Holding Register (3) should default to 1 (Read/Write)
        rows = [{'Name': 'HoldingVar', 'Address': '200', 'Type': 'U16', 'RegisterType': 'Holding Register', 'Action': ''}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

        # Discrete Input (2) should default to 4 (Read Only)
        rows = [{'Name': 'DiscVar', 'Address': '300', 'Type': 'U16', 'RegisterType': 'Discrete Input', 'Action': ''}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4')

        # Coil (1) should default to 1 (Read/Write)
        rows = [{'Name': 'CoilVar', 'Address': '400', 'Type': 'U16', 'RegisterType': 'Coil', 'Action': ''}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

if __name__ == "__main__":
    unittest.main()
