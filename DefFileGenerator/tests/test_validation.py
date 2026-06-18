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

    def tearDown(self):
        self.temp_dir.cleanup()

    def create_def_file(self, rows, manufacturer="Test", model="Model"):
        filepath = os.path.join(self.temp_dir.name, "test_def.csv")
        header = ["modbusRTU", "Inverter", manufacturer, model, "", "", "", "", "", "", ""]
        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(header)
            for i, row in enumerate(rows, start=1):
                # row format: [Info1, Info2, Info3, Info4, Name, Tag, CoefA, CoefB, Unit, Action]
                full_row = [str(i)] + row
                writer.writerow(full_row)
        return filepath

    def test_validate_address_range(self):
        # Valid addresses
        self.assertTrue(self.generator.validate_address('0', 'U16'))
        self.assertTrue(self.generator.validate_address('65535', 'U16'))
        self.assertTrue(self.generator.validate_address('0x0', 'U16'))
        self.assertTrue(self.generator.validate_address('0xFFFF', 'U16'))

        # Invalid addresses (out of range)
        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(self.generator.validate_address('65536', 'U16'))
            self.assertTrue(any("out of standard Modbus range" in m for m in log.output))

        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(self.generator.validate_address('-1', 'U16'))
            self.assertTrue(any("out of standard Modbus range" in m for m in log.output))

    def test_validate_csv_valid(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']
        ]
        filepath = self.create_def_file(rows)
        self.assertTrue(self.generator.validate_csv(filepath))

    def test_validate_csv_duplicate_tag(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'A', '4'] # Duplicate Tag
        ]
        filepath = self.create_def_file(rows)
        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(self.generator.validate_csv(filepath))
            self.assertTrue(any("Duplicate Tag 'tag1'" in m for m in log.output))

    def test_validate_csv_overlap(self):
        rows = [
            ['3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'], # Uses 100, 101
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']  # Overlaps 101
        ]
        filepath = self.create_def_file(rows)
        with self.assertLogs(level='WARNING') as log:
            # validate_csv returns True for overlaps (as they are currently warnings)
            self.assertTrue(self.generator.validate_csv(filepath))
            self.assertTrue(any("Address overlap detected" in m for m in log.output))

    def test_validate_csv_invalid_address(self):
        rows = [
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'] # Out of range
        ]
        filepath = self.create_def_file(rows)
        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(self.generator.validate_csv(filepath))
            self.assertTrue(any("Invalid address or type" in m for m in log.output))

    def test_summary_logging(self):
        rows = [
            {'Info1': '3', 'Info2': '100', 'Info3': 'U16', 'Info4': '', 'Name': 'V1', 'Tag': 't1', 'CoefA': '1.0', 'CoefB': '0.0', 'Unit': 'V', 'Action': '4'},
            {'Info1': '4', 'Info2': '200', 'Info3': 'U16', 'Info4': '', 'Name': 'V2', 'Tag': 't2', 'CoefA': '1.0', 'CoefB': '0.0', 'Unit': 'A', 'Action': '4'}
        ]
        output = os.path.join(self.temp_dir.name, "summary.csv")
        with self.assertLogs(level='INFO') as log:
            Generator.write_output_csv(output, rows, "MFG", "MDL")
            self.assertTrue(any("Generated registers: Holding: 1, Input: 1" in m for m in log.output))

if __name__ == '__main__':
    unittest.main()
