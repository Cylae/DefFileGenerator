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
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)

    def test_validate_address_range(self):
        # Valid range 0-65535
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        self.assertTrue(self.generator.validate_address("0x0", "U16"))
        self.assertTrue(self.generator.validate_address("0xFFFF", "U16"))

        # Out of range
        self.assertFalse(self.generator.validate_address("65536", "U16"))
        self.assertFalse(self.generator.validate_address("-1", "U16"))
        self.assertFalse(self.generator.validate_address("0x10000", "U16"))

    def test_validate_compound_address_range(self):
        # Valid compound
        self.assertTrue(self.generator.validate_address("100_20", "STRING"))
        self.assertTrue(self.generator.validate_address("0x100_10", "STR10"))

        # Out of range compound
        self.assertFalse(self.generator.validate_address("70000_20", "STRING"))

    def test_validate_csv_logic(self):
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, newline='') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U16', '', 'Var1', 'var1', '1.0', '0.0', 'V', '4'])
            tmp_path = f.name

        try:
            self.assertTrue(self.generator.validate_csv(tmp_path))

            # Test invalid data type
            with open(tmp_path, 'a', newline='') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow(['2', '3', '101', 'INVALID', '', 'Var2', 'var2', '1.0', '0.0', 'A', '4'])
            self.assertFalse(self.generator.validate_csv(tmp_path))
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def test_validate_csv_duplicate_tags(self):
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, newline='') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U16', '', 'Var1', 'dup_tag', '1.0', '0.0', 'V', '4'])
            writer.writerow(['2', '3', '101', 'U16', '', 'Var2', 'dup_tag', '1.0', '0.0', 'A', '4'])
            tmp_path = f.name

        try:
            self.assertFalse(self.generator.validate_csv(tmp_path))
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

if __name__ == '__main__':
    unittest.main()
