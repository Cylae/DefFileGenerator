import unittest
import os
import logging
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_file = "test_validate.csv"
        logging.basicConfig(level=logging.INFO)

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def write_test_csv(self, rows):
        with open(self.test_file, 'w', encoding='utf-8') as f:
            f.write("modbusRTU;Inverter;TestMfg;TestModel;;;;;;;\n")
            for i, row in enumerate(rows, start=1):
                f.write(f"{i};" + ";".join(row) + "\n")

    def test_validate_success(self):
        self.write_test_csv([
            ["3", "100", "U16", "", "Var1", "tag1", "1.0", "0.0", "V", "4"],
            ["3", "101", "U16", "", "Var2", "tag2", "1.0", "0.0", "V", "4"]
        ])
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_validate_duplicate_tag(self):
        self.write_test_csv([
            ["3", "100", "U16", "", "Var1", "tag1", "1.0", "0.0", "V", "4"],
            ["3", "101", "U16", "", "Var2", "tag1", "1.0", "0.0", "V", "4"]
        ])
        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Duplicate Tag 'tag1'" in m for m in log.output))

    def test_validate_address_overlap(self):
        self.write_test_csv([
            ["3", "100", "U32", "", "Var1", "tag1", "1.0", "0.0", "V", "4"],
            ["3", "101", "U16", "", "Var2", "tag2", "1.0", "0.0", "V", "4"]
        ])
        # Overlap is a warning in the current implementation, but validate_csv should still run
        # Address 100 for U32 uses 100 and 101. Address 101 overlaps.
        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("overlap detected" in m.lower() for m in log.output))

    def test_validate_out_of_range_address(self):
        self.write_test_csv([
            ["3", "70000", "U16", "", "Var1", "tag1", "1.0", "0.0", "V", "4"]
        ])
        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("out of standard Modbus range" in m for m in log.output))

    def test_validate_invalid_header(self):
        with open(self.test_file, 'w') as f:
            f.write("invalid;header\n")
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_validate_insufficient_columns(self):
        with open(self.test_file, 'w') as f:
            f.write("modbusRTU;Inverter;TestMfg;TestModel;;;;;;;\n")
            f.write("1;3;100;U16;MissingRest\n")
        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("insufficient columns" in m for m in log.output))

if __name__ == "__main__":
    unittest.main()
