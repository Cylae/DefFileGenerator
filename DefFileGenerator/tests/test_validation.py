import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_csv = "test_validation.csv"

    def tearDown(self):
        if os.path.exists(self.test_csv):
            os.remove(self.test_csv)

    def create_csv(self, rows):
        fieldnames = ['Name', 'Address', 'Type']
        with open(self.test_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)

    def test_validate_address_range(self):
        # Valid range
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        self.assertTrue(self.generator.validate_address("0xFFFF", "U16"))

        # Out of range
        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(self.generator.validate_address("65536", "U16"))
            self.assertTrue(any("out of range" in m for m in log.output))

        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(self.generator.validate_address("-1", "U16"))
            self.assertTrue(any("out of range" in m for m in log.output))

    def test_validate_csv_success(self):
        self.create_csv([
            {'Name': 'Var1', 'Address': '100', 'Type': 'U16'},
            {'Name': 'Var2', 'Address': '0x200', 'Type': 'I32'}
        ])
        self.assertTrue(self.generator.validate_csv(self.test_csv))

    def test_validate_csv_failure(self):
        self.create_csv([
            {'Name': 'Var1', 'Address': '70000', 'Type': 'U16'}, # Out of range
        ])
        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(self.generator.validate_csv(self.test_csv))
            self.assertTrue(any("Invalid address or range" in m for m in log.output))

    def test_validate_csv_missing_column(self):
        with open(self.test_csv, 'w', newline='', encoding='utf-8') as f:
            f.write("Name,Type\nVar1,U16\n")
        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(self.generator.validate_csv(self.test_csv))
            self.assertTrue(any("Missing 'Address' column" in m for m in log.output))

    def test_intelligent_action_defaulting(self):
        rows = [
            {'Name': 'Holding', 'Address': '100', 'RegisterType': 'Holding', 'Type': 'U16'}, # Default 1
            {'Name': 'Input', 'Address': '200', 'RegisterType': 'Input', 'Type': 'U16'},     # Default 4
            {'Name': 'Discrete', 'Address': '300_0_1', 'RegisterType': 'Discrete', 'Type': 'BITS'}, # Default 4
            {'Name': 'Coil', 'Address': '400', 'RegisterType': 'Coil', 'Type': 'U16'},      # Default 1
        ]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')
        self.assertEqual(processed[1]['Action'], '4')
        self.assertEqual(processed[2]['Action'], '4')
        self.assertEqual(processed[3]['Action'], '1')

    def test_normalize_type_string_n(self):
        self.assertEqual(self.generator.normalize_type("string 20"), "STR20")
        self.assertEqual(self.generator.normalize_type("string 32_WB"), "STR32_WB")

if __name__ == '__main__':
    unittest.main()
