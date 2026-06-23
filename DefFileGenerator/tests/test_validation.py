import unittest
import os
import csv
import tempfile
import logging
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        # Disable logging for tests to keep output clean
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)

    def test_validate_address_range(self):
        # Valid addresses
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        self.assertTrue(self.generator.validate_address("0x0", "U16"))
        self.assertTrue(self.generator.validate_address("0xFFFF", "U16"))

        # Invalid addresses (out of range)
        self.assertFalse(self.generator.validate_address("65536", "U16"))
        self.assertFalse(self.generator.validate_address("-1", "U16"))
        self.assertFalse(self.generator.validate_address("0x10000", "U16"))

    def test_validate_address_formats(self):
        # STR<n> synonyms
        self.assertTrue(self.generator.validate_address("100_10", "STR20"))
        self.assertTrue(self.generator.validate_address("100_10", "STRING"))

        # BITS
        self.assertTrue(self.generator.validate_address("100_0_1", "BITS"))
        self.assertFalse(self.generator.validate_address("100", "BITS"))

    def test_action_defaulting(self):
        rows = [
            {'Name': 'Hold', 'Address': '1', 'RegisterType': 'Holding', 'Type': 'U16', 'Action': ''},
            {'Name': 'In', 'Address': '2', 'RegisterType': 'Input', 'Type': 'U16', 'Action': ''},
            {'Name': 'Coil', 'Address': '3', 'RegisterType': 'Coil', 'Type': 'U16', 'Action': ''},
            {'Name': 'Disc', 'Address': '4', 'RegisterType': 'Discrete Input', 'Type': 'U16', 'Action': ''},
        ]
        processed = list(self.generator.process_rows(rows))

        # Holding defaults to Read/Write (1)
        self.assertEqual(processed[0]['Action'], '1')
        # Input defaults to Read Only (4)
        self.assertEqual(processed[1]['Action'], '4')
        # Coil defaults to Read/Write (1)
        self.assertEqual(processed[2]['Action'], '1')
        # Discrete defaults to Read Only (4)
        self.assertEqual(processed[3]['Action'], '4')

    def test_validate_csv_duplicates(self):
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Mfg', 'Model', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])
            writer.writerow(['2', '3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4']) # Duplicate Tag
            temp_path = f.name

        try:
            self.assertFalse(self.generator.validate_csv(temp_path))
        finally:
            os.unlink(temp_path)

    def test_validate_csv_overlaps(self):
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Mfg', 'Model', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']) # Uses 100, 101
            writer.writerow(['2', '3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']) # Overlaps 101
            temp_path = f.name

        try:
            # Overlaps are currently logged as warnings but don't fail validate_csv unless fatal
            # Based on my implementation, validate_csv returns valid=True if no fatal errors (like duplicate tags or bad addresses)
            # Address overlaps are warnings.
            self.assertTrue(self.generator.validate_csv(temp_path))
        finally:
            os.unlink(temp_path)

    def test_validate_csv_bad_address(self):
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Mfg', 'Model', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']) # Bad address
            temp_path = f.name

        try:
            self.assertFalse(self.generator.validate_csv(temp_path))
        finally:
            os.unlink(temp_path)

if __name__ == '__main__':
    unittest.main()
