import unittest
import os
import csv
import tempfile
import logging
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.test_dir.cleanup()

    def test_validate_address_range(self):
        # Valid addresses
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        self.assertTrue(self.generator.validate_address("0x0", "U16"))
        self.assertTrue(self.generator.validate_address("0xFFFF", "U16"))
        self.assertTrue(self.generator.validate_address("40001", "U16"))

        # Invalid addresses (out of range)
        with self.assertLogs(level='WARNING') as cm:
            self.assertFalse(self.generator.validate_address("65536", "U16"))
            self.assertIn("out of standard Modbus range", cm.output[0])

        self.assertFalse(self.generator.validate_address("-1", "U16"))
        self.assertFalse(self.generator.validate_address("0x10000", "U16"))

    def test_validate_address_str_synonym(self):
        # STR<n> synonyms should be recognized during validation
        self.assertTrue(self.generator.validate_address("30001_10", "STR20"))
        self.assertTrue(self.generator.validate_address("30001_10", "STRING"))
        self.assertFalse(self.generator.validate_address("30001", "STR20")) # Compound addr expected for strings

    def test_validate_csv(self):
        def_file = os.path.join(self.test_dir.name, "test_def.csv")

        # Create a valid definition file
        with open(def_file, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '30001', 'U16', '', 'Var1', 'var1', '1.0', '0.0', 'V', '4'])

        self.assertTrue(self.generator.validate_csv(def_file))

        # Invalid: Address out of range
        with open(def_file, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '70000', 'U16', '', 'Var1', 'var1', '1.0', '0.0', 'V', '4'])

        self.assertFalse(self.generator.validate_csv(def_file))

    def test_action_defaulting(self):
        rows = [
            {'Name': 'Holding', 'Address': '30001', 'Type': 'U16', 'RegisterType': 'Holding Register'},
            {'Name': 'Input', 'Address': '40001', 'Type': 'U16', 'RegisterType': 'Input Register'},
            {'Name': 'Coil', 'Address': '1', 'Type': 'U16', 'RegisterType': 'Coil'},
            {'Name': 'Discrete', 'Address': '10001', 'Type': 'U16', 'RegisterType': 'Discrete Input'},
        ]

        processed = list(self.generator.process_rows(rows))

        # Holding defaults to 1 (RW)
        self.assertEqual(processed[0]['Action'], '1')
        # Input defaults to 4 (RO)
        self.assertEqual(processed[1]['Action'], '4')
        # Coil defaults to 1 (RW)
        self.assertEqual(processed[2]['Action'], '1')
        # Discrete defaults to 4 (RO)
        self.assertEqual(processed[3]['Action'], '4')

if __name__ == '__main__':
    unittest.main()
