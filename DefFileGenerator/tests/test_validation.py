import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidationLogic(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_def = "test_validate_def.csv"

    def tearDown(self):
        if os.path.exists(self.test_def):
            os.remove(self.test_def)

    def write_def(self, rows, manufacturer="TestMfg", model="TestModel"):
        with open(self.test_def, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', manufacturer, model, '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + list(row))

    def test_validate_csv_success(self):
        rows = [
            ['3', '100', 'U16', '', 'Voltage', 'voltage', '1', '0', 'V', '4'],
            ['3', '101', 'U16', '', 'Current', 'current', '1', '0', 'A', '4']
        ]
        self.write_def(rows)
        self.assertTrue(self.generator.validate_csv(self.test_def))

    def test_validate_csv_duplicate_tag(self):
        rows = [
            ['3', '100', 'U16', '', 'Voltage', 'voltage', '1', '0', 'V', '4'],
            ['3', '101', 'U16', '', 'Current', 'voltage', '1', '0', 'A', '4']
        ]
        self.write_def(rows)
        with self.assertLogs(level='ERROR') as cm:
            self.assertFalse(self.generator.validate_csv(self.test_def))
            self.assertTrue(any("FATAL: Duplicate Tag 'voltage'" in line for line in cm.output))

    def test_validate_csv_address_overlap(self):
        rows = [
            ['3', '100', 'U32', '', 'Power', 'power', '1', '0', 'W', '4'],
            ['3', '101', 'U16', '', 'Overlap', 'overlap', '1', '0', '', '4']
        ]
        self.write_def(rows)
        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_csv(self.test_def)) # Overlap is a warning, not fatal for validity currently in logic, though it might be desired to be false. Let's check def_gen.py logic.
            # In my implementation of validate_csv, overlaps only log warnings.
            self.assertTrue(any("Address overlap at 101" in line for line in cm.output))

    def test_validate_address_range(self):
        # 0-65535 is valid
        self.assertTrue(Generator.validate_address("0", "U16"))
        self.assertTrue(Generator.validate_address("65535", "U16"))

        # Out of range should log warning but still return True for format validity
        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(Generator.validate_address("65536", "U16"))
            self.assertTrue(any("out of standard Modbus range" in line for line in cm.output))

    def test_validate_csv_insufficient_columns(self):
        with open(self.test_def, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', ''])
            writer.writerow(['1', '3', '100']) # Too short

        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_csv(self.test_def))
            self.assertTrue(any("Insufficient columns" in line for line in cm.output))

if __name__ == "__main__":
    unittest.main()
