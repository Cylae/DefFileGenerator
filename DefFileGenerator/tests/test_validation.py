import unittest
import logging
import os
import csv
import tempfile
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.test_dir.cleanup()

    def test_validate_address_range(self):
        # Standard Modbus range is 0-65535
        self.assertTrue(self.generator.validate_address('0', 'U16'))
        self.assertTrue(self.generator.validate_address('65535', 'U16'))

        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(self.generator.validate_address('65536', 'U16'))
            self.assertTrue(any("outside standard Modbus range" in m for m in log.output))

        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(self.generator.validate_address('-1', 'U16'))
            self.assertTrue(any("outside standard Modbus range" in m for m in log.output))

    def test_validate_address_str_n(self):
        # STR20 should accept both simple and compound addresses
        self.assertTrue(self.generator.validate_address('100', 'STR20'))
        self.assertTrue(self.generator.validate_address('100_20', 'STR20'))
        # But not invalid formats
        self.assertFalse(self.generator.validate_address('100_20_5', 'STR20'))

    def test_validate_csv_success(self):
        csv_path = os.path.join(self.test_dir.name, 'valid.csv')
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=['Name', 'Address', 'Type', 'RegisterType'])
            writer.writeheader()
            writer.writerow({'Name': 'Var1', 'Address': '100', 'Type': 'U16', 'RegisterType': 'Holding'})
            writer.writerow({'Name': 'Var2', 'Address': '200', 'Type': 'STR20', 'RegisterType': 'Holding'})

        self.assertTrue(self.generator.validate_csv(csv_path))

    def test_validate_csv_failure_type(self):
        csv_path = os.path.join(self.test_dir.name, 'invalid_type.csv')
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=['Name', 'Address', 'Type', 'RegisterType'])
            writer.writeheader()
            writer.writerow({'Name': 'Var1', 'Address': '100', 'Type': 'INVALID', 'RegisterType': 'Holding'})

        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(self.generator.validate_csv(csv_path))
            self.assertTrue(any("Invalid data type" in m for m in log.output))

    def test_validate_csv_failure_address(self):
        csv_path = os.path.join(self.test_dir.name, 'invalid_addr.csv')
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=['Name', 'Address', 'Type', 'RegisterType'])
            writer.writeheader()
            writer.writerow({'Name': 'Var1', 'Address': '70000', 'Type': 'U16', 'RegisterType': 'Holding'})

        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(self.generator.validate_csv(csv_path))
            self.assertTrue(any("Invalid address" in m for m in log.output))

    def test_action_defaults(self):
        # Input (4) and Discrete (2) default to RO (4)
        # Holding (3) and Coils (1) default to RW (1)
        rows = [
            {'Name': 'In', 'Address': '1', 'Type': 'U16', 'RegisterType': 'Input Register', 'Action': ''},
            {'Name': 'Disc', 'Address': '2', 'Type': 'U16', 'RegisterType': 'Discrete Input', 'Action': ''},
            {'Name': 'Hold', 'Address': '3', 'Type': 'U16', 'RegisterType': 'Holding Register', 'Action': ''},
            {'Name': 'Coil', 'Address': '4', 'Type': 'U16', 'RegisterType': 'Coil', 'Action': ''},
        ]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4') # Input -> RO
        self.assertEqual(processed[1]['Action'], '4') # Discrete -> RO
        self.assertEqual(processed[2]['Action'], '1') # Holding -> RW
        self.assertEqual(processed[3]['Action'], '1') # Coil -> RW

if __name__ == '__main__':
    unittest.main()
