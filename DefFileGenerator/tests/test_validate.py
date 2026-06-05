import unittest
import os
import csv
import tempfile
import logging
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_dir = tempfile.TemporaryDirectory()
        self.test_file = os.path.join(self.test_dir.name, "test_def.csv")
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)
        self.test_dir.cleanup()

    def create_def_file(self, rows, header=None):
        if header is None:
            header = ["modbusRTU", "Inverter", "TestMFG", "TestModel", "", "", "", "", "", "", ""]
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(header)
            for i, row in enumerate(rows, start=1):
                # Format: Index, Info1, Info2, Info3, Info4, Name, Tag, CoefA, CoefB, Unit, Action
                writer.writerow([str(i)] + row)

    def test_validate_success(self):
        rows = [
            ["3", "40001", "U16", "", "Voltage", "voltage", "1.0", "0.0", "V", "4"],
            ["3", "40002", "U16", "", "Current", "current", "0.1", "0.0", "A", "4"]
        ]
        self.create_def_file(rows)
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_validate_duplicate_tag(self):
        rows = [
            ["3", "40001", "U16", "", "Voltage", "v_tag", "1.0", "0.0", "V", "4"],
            ["3", "40002", "U16", "", "Current", "v_tag", "0.1", "0.0", "A", "4"]
        ]
        self.create_def_file(rows)
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_validate_address_overlap(self):
        rows = [
            ["3", "40001", "U32", "", "Power", "power", "1.0", "0.0", "W", "4"],
            ["3", "40002", "U16", "", "Voltage", "voltage", "1.0", "0.0", "V", "4"]
        ]
        self.create_def_file(rows)
        # _check_address_overlap logs a warning but validate_csv should still return True
        # unless it's a fatal error. Duplicate Tag is fatal.
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_validate_invalid_header(self):
        self.create_def_file([], header=["Invalid"])
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_validate_out_of_range_address(self):
        rows = [
            ["3", "70000", "U16", "", "OutRange", "out_range", "1.0", "0.0", "", "4"]
        ]
        self.create_def_file(rows)
        # Range check is a warning
        self.assertTrue(self.generator.validate_csv(self.test_file))

if __name__ == '__main__':
    unittest.main()
