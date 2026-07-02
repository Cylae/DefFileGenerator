import unittest
import os
import tempfile
import csv
import logging
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        # Suppress logging during tests
        logging.getLogger().setLevel(logging.ERROR)

    def test_validate_address_range(self):
        # Valid addresses
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        self.assertTrue(self.generator.validate_address("0x0", "U16"))
        self.assertTrue(self.generator.validate_address("0xFFFF", "U16"))
        self.assertTrue(self.generator.validate_address("100h", "U16"))

        # Invalid addresses (out of range)
        self.assertFalse(self.generator.validate_address("65536", "U16"))
        self.assertFalse(self.generator.validate_address("-1", "U16"))
        self.assertFalse(self.generator.validate_address("0x10000", "U16"))

        # Compound addresses
        self.assertTrue(self.generator.validate_address("100_20", "STRING"))
        self.assertTrue(self.generator.validate_address("100_20", "STR20"))
        self.assertTrue(self.generator.validate_address("65535_20", "STRING"))
        self.assertFalse(self.generator.validate_address("65536_20", "STRING"))

        self.assertTrue(self.generator.validate_address("100_0_1", "BITS"))
        self.assertFalse(self.generator.validate_address("65536_0_1", "BITS"))

    def test_action_defaulting(self):
        # Input/Discrete should default to '4' (Read Only)
        rows = [
            {'Name': 'In1', 'Address': '100', 'Type': 'U16', 'RegisterType': 'Input Register', 'Action': ''},
            {'Name': 'In2', 'Address': '101', 'Type': 'U16', 'RegisterType': 'input', 'Action': ''},
            {'Name': 'Disc1', 'Address': '102', 'Type': 'U16', 'RegisterType': 'Discrete Input', 'Action': ''},
        ]
        processed = list(self.generator.process_rows(rows))
        for p in processed:
            self.assertEqual(p['Action'], '4', f"Failed for {p['Name']}")

        # Holding/Coils should default to '1' (Read/Write)
        rows = [
            {'Name': 'Hold1', 'Address': '200', 'Type': 'U16', 'RegisterType': 'Holding Register', 'Action': ''},
            {'Name': 'Coil1', 'Address': '201', 'Type': 'U16', 'RegisterType': 'Coil', 'Action': ''},
        ]
        processed = list(self.generator.process_rows(rows))
        for p in processed:
            self.assertEqual(p['Action'], '1', f"Failed for {p['Name']}")

    def test_validate_csv_file(self):
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as tf:
            writer = csv.writer(tf, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Mfg', 'Model', '', '', '', '', '', '', ''])
            # Valid row
            writer.writerow(['1', '3', '100', 'U16', '', 'Var1', 'var1', '1.0', '0.0', 'V', '4'])
            # Invalid type row
            writer.writerow(['2', '3', '101', 'INVALID', '', 'Var2', 'var2', '1.0', '0.0', 'V', '4'])
            temp_name = tf.name

        try:
            # Should fail because of Var2
            self.assertFalse(self.generator.validate_csv(temp_name))

            # Create a valid one
            with open(temp_name, 'w', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow(['modbusRTU', 'Inverter', 'Mfg', 'Model', '', '', '', '', '', '', ''])
                writer.writerow(['1', '3', '100', 'U16', '', 'Var1', 'var1', '1.0', '0.0', 'V', '4'])

            self.assertTrue(self.generator.validate_csv(temp_name))
        finally:
            if os.path.exists(temp_name):
                os.remove(temp_name)

    def test_validate_csv_compound_address(self):
         with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as tf:
            writer = csv.writer(tf, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Mfg', 'Model', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '30030_20', 'STRING', '', 'Str1', 'str1', '1.0', '0.0', '', '4'])
            temp_name = tf.name
         try:
            self.assertTrue(self.generator.validate_csv(temp_name))
         finally:
            if os.path.exists(temp_name):
                os.remove(temp_name)

if __name__ == '__main__':
    unittest.main()
