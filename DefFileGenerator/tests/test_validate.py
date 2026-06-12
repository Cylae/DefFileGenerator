import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.valid_file = "test_valid.csv"
        self.invalid_file = "test_invalid.csv"

        # Valid file
        with open(self.valid_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '40001', 'U16', '', 'Var1', 'var1', '1.0', '0.0', 'V', '1'])
            writer.writerow(['2', '3', '40002', 'U32', '', 'Var2', 'var2', '1.0', '0.0', 'A', '1'])

    def tearDown(self):
        for f in [self.valid_file, self.invalid_file]:
            if os.path.exists(f):
                os.remove(f)

    def test_validate_valid(self):
        self.assertTrue(self.generator.validate_csv(self.valid_file))

    def test_validate_duplicate_tag(self):
        with open(self.invalid_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '40001', 'U16', '', 'Var1', 'dup', '1.0', '0.0', 'V', '1'])
            writer.writerow(['2', '3', '40002', 'U16', '', 'Var2', 'dup', '1.0', '0.0', 'A', '1'])

        with self.assertLogs(level='ERROR') as cm:
            self.assertFalse(self.generator.validate_csv(self.invalid_file))
            self.assertTrue(any("Duplicate Tag 'dup'" in output for output in cm.output))

    def test_validate_overlap(self):
        with open(self.invalid_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '40001', 'U32', '', 'Var1', 'v1', '1.0', '0.0', 'V', '1'])
            writer.writerow(['2', '3', '40002', 'U16', '', 'Var2', 'v2', '1.0', '0.0', 'A', '1'])

        with self.assertLogs(level='WARNING') as cm:
            # Overlap should log warning but not necessarily return False unless we want it to.
            # Current implementation return False ONLY on duplicate tags (Fatal).
            # Let's check if it returns True but logs warning.
            self.assertTrue(self.generator.validate_csv(self.invalid_file))
            self.assertTrue(any("Address overlap detected" in output for output in cm.output))

    def test_validate_out_of_range(self):
        with open(self.invalid_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '70000', 'U16', '', 'Var1', 'v1', '1.0', '0.0', 'V', '1'])

        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_csv(self.invalid_file))
            self.assertTrue(any("Address 70000 is out of standard Modbus range" in output for output in cm.output))

if __name__ == "__main__":
    unittest.main()
