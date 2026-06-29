import unittest
import os
import csv
import tempfile
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        # Suppress logging during tests
        logging.getLogger().setLevel(logging.ERROR)

    def test_validate_address_range(self):
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        self.assertFalse(self.generator.validate_address("65536", "U16"))
        self.assertFalse(self.generator.validate_address("-1", "U16"))

    def test_validate_address_hex_range(self):
        self.assertTrue(self.generator.validate_address("0x0", "U16"))
        self.assertTrue(self.generator.validate_address("0xFFFF", "U16"))
        self.assertFalse(self.generator.validate_address("0x10000", "U16"))

    def test_validate_type_synonyms(self):
        self.assertTrue(self.generator.validate_type("STR20"))
        self.assertEqual(self.generator.normalize_type("string 20"), "STR20")
        self.assertEqual(self.generator.normalize_type("string"), "STRING")

    def test_action_defaulting(self):
        rows = [
            {'Name': 'Reg1', 'Address': '1', 'RegisterType': 'Input Register', 'Type': 'U16'},
            {'Name': 'Reg2', 'Address': '2', 'RegisterType': 'Holding Register', 'Type': 'U16'},
            {'Name': 'Reg3', 'Address': '3', 'RegisterType': 'Discrete Input', 'Type': 'U16'},
            {'Name': 'Reg4', 'Address': '4', 'RegisterType': 'Coil', 'Type': 'U16'},
        ]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4') # Input -> Read Only
        self.assertEqual(processed[1]['Action'], '1') # Holding -> Read/Write
        self.assertEqual(processed[2]['Action'], '4') # Discrete -> Read Only
        self.assertEqual(processed[3]['Action'], '1') # Coil -> Read/Write

    def test_validate_csv_method(self):
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as f:
            writer = csv.writer(f)
            writer.writerow(['Name', 'Address', 'RegisterType', 'Type'])
            writer.writerow(['Valid', '100', 'Holding Register', 'U16'])
            writer.writerow(['InvalidAddr', '70000', 'Holding Register', 'U16'])
            temp_path = f.name

        try:
            # Should be False because of invalid address 70000
            self.assertFalse(self.generator.validate_csv(temp_path))

            # Now test a valid one
            with open(temp_path, 'w') as f:
                writer = csv.writer(f)
                writer.writerow(['Name', 'Address', 'RegisterType', 'Type'])
                writer.writerow(['Valid', '100', 'Holding Register', 'U16'])
            self.assertTrue(self.generator.validate_csv(temp_path))
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)

if __name__ == '__main__':
    unittest.main()
