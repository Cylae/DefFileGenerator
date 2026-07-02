import unittest
import os
import csv
import logging
import tempfile
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        # Suppress logging during tests
        logging.getLogger().setLevel(logging.ERROR)

    def test_validate_address_range(self):
        # Valid addresses
        self.assertTrue(Generator.validate_address("0", "U16"))
        self.assertTrue(Generator.validate_address("65535", "U16"))
        self.assertTrue(Generator.validate_address("0x0", "U16"))
        self.assertTrue(Generator.validate_address("0xFFFF", "U16"))
        self.assertTrue(Generator.validate_address("30001_10", "STRING"))
        self.assertTrue(Generator.validate_address("40001_0_1", "BITS"))

        # Invalid addresses (out of range)
        self.assertFalse(Generator.validate_address("65536", "U16"))
        self.assertFalse(Generator.validate_address("-1", "U16"))
        self.assertFalse(Generator.validate_address("0x10000", "U16"))
        self.assertFalse(Generator.validate_address("70000_10", "STRING"))

    def test_validate_address_synonyms(self):
        # Test STR20 synonym in validate_address
        self.assertTrue(Generator.validate_address("30030_20", "STR20"))
        self.assertFalse(Generator.validate_address("30030", "STR20"))

    def test_validate_csv(self):
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.csv') as tmp:
            tmp_path = tmp.name
            writer = csv.writer(tmp, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U16', '', 'V1', 'tag1', '1.0', '0.0', 'V', '4'])
            writer.writerow(['2', '3', '70000', 'U16', '', 'V2', 'tag2', '1.0', '0.0', 'V', '4']) # Invalid address
            tmp.close()

            try:
                self.assertFalse(Generator.validate_csv(tmp_path))
            finally:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)

    def test_validate_csv_insufficient_columns(self):
         with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.csv') as tmp:
            tmp_path = tmp.name
            writer = csv.writer(tmp, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U16']) # Insufficient columns
            tmp.close()

            try:
                # Should skip row and still return True (or at least not crash)
                # In our implementation, it logs a warning and continues.
                # If all valid rows are ok, it returns success.
                self.assertTrue(Generator.validate_csv(tmp_path))
            finally:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)

    def test_action_defaulting(self):
        rows = [
            {'Name': 'InputVar', 'Address': '100', 'RegisterType': 'Input Register', 'Type': 'U16'},
            {'Name': 'HoldingVar', 'Address': '200', 'RegisterType': 'Holding Register', 'Type': 'U16'}
        ]
        processed = list(self.generator.process_rows(rows))

        self.assertEqual(processed[0]['Action'], '4') # Input defaults to Read-Only
        self.assertEqual(processed[1]['Action'], '1') # Holding defaults to Read/Write

if __name__ == '__main__':
    unittest.main()
