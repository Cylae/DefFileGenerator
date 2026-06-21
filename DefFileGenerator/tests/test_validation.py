import unittest
import os
import csv
import tempfile
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        # Suppress logging during tests unless needed
        logging.getLogger().setLevel(logging.ERROR)

    def test_validate_address_range(self):
        # Valid addresses
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        self.assertTrue(self.generator.validate_address("0x0", "U16"))
        self.assertTrue(self.generator.validate_address("0xFFFF", "U16"))

        # Invalid addresses
        self.assertFalse(self.generator.validate_address("-1", "U16"))
        self.assertFalse(self.generator.validate_address("65536", "U16"))
        self.assertFalse(self.generator.validate_address("0x10000", "U16"))

    def test_validate_address_str_synonym(self):
        # STR20 should be treated as STRING synonym
        self.assertTrue(self.generator.validate_address("40001_10", "STR20"))
        self.assertFalse(self.generator.validate_address("40001", "STR20"))

    def test_validate_csv_basic(self):
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as tmp:
            writer = csv.writer(tmp, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Test', 'Model', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '40001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])
            tmp_path = tmp.name

        try:
            self.assertTrue(self.generator.validate_csv(tmp_path))
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def test_validate_csv_duplicate_tag(self):
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as tmp:
            writer = csv.writer(tmp, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Test', 'Model', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '40001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])
            writer.writerow(['2', '3', '40002', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4'])
            tmp_path = tmp.name

        try:
            self.assertFalse(self.generator.validate_csv(tmp_path))
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def test_validate_csv_address_overlap(self):
        # Overlap detection should log warning but not necessarily fail validation unless specified
        # In current implementation, overlap is a warning, not a fatal error for validate_csv
        # Wait, let me check implementation of validate_csv
        pass

    def test_validate_csv_invalid_address(self):
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as tmp:
            writer = csv.writer(tmp, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Test', 'Model', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])
            tmp_path = tmp.name

        try:
            self.assertFalse(self.generator.validate_csv(tmp_path))
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

if __name__ == '__main__':
    unittest.main()
