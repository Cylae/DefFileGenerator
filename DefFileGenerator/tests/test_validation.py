import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.def_file = "test_validation.csv"
        self.log_stream = unittest.mock.MagicMock()

    def tearDown(self):
        if os.path.exists(self.def_file):
            os.remove(self.def_file)

    def write_def_file(self, rows):
        with open(self.def_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Test Mfg', 'Test Model', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + list(row))

    def test_validate_csv_success(self):
        rows = [
            ('3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '1'),
            ('4', '200', 'F32', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4')
        ]
        self.write_def_file(rows)
        self.assertTrue(Generator.validate_csv(self.def_file))

    def test_validate_csv_duplicate_tag(self):
        rows = [
            ('3', '100', 'U16', '', 'Var1', 'dup', '1.0', '0.0', 'V', '1'),
            ('3', '101', 'U16', '', 'Var2', 'dup', '1.0', '0.0', 'A', '1')
        ]
        self.write_def_file(rows)
        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(Generator.validate_csv(self.def_file))
            self.assertTrue(any("Duplicate Tag 'dup'" in m for m in log.output))

    def test_validate_csv_address_overlap(self):
        rows = [
            ('3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '1'),
            ('3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '1')
        ]
        self.write_def_file(rows)
        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(Generator.validate_csv(self.def_file))
            self.assertTrue(any("Address overlap detected" in m for m in log.output))

    def test_validate_address_range(self):
        # Valid
        self.assertTrue(Generator.validate_address("0", "U16"))
        self.assertTrue(Generator.validate_address("65535", "U16"))
        self.assertTrue(Generator.validate_address("0x0", "U16"))
        self.assertTrue(Generator.validate_address("0xFFFF", "U16"))

        # Invalid
        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(Generator.validate_address("65536", "U16"))
            self.assertTrue(any("out of standard Modbus range" in m for m in log.output))

        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(Generator.validate_address("-1", "U16"))
            self.assertTrue(any("out of standard Modbus range" in m for m in log.output))

    def test_validate_csv_insufficient_cols(self):
        with open(self.def_file, 'w', newline='', encoding='utf-8') as f:
            f.write("modbusRTU;Inverter;Mfg;Model;;;;;;;\n")
            f.write("1;3;100;U16;Name;tag\n") # Only 6 cols
        with self.assertLogs(level='WARNING') as log:
            # It should still be valid if no other errors, but log a warning
            self.assertTrue(Generator.validate_csv(self.def_file))
            self.assertTrue(any("Insufficient columns" in m for m in log.output))

if __name__ == "__main__":
    unittest.main()
