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
        # Valid addresses
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        self.assertTrue(self.generator.validate_address("0x0", "U16"))
        self.assertTrue(self.generator.validate_address("0xFFFF", "U16"))

        # Invalid addresses (out of range)
        self.assertFalse(self.generator.validate_address("65536", "U16"))
        self.assertFalse(self.generator.validate_address("-1", "U16"))
        self.assertFalse(self.generator.validate_address("0x10000", "U16"))

    def test_validate_address_compound(self):
        # STR<n> synonyms
        self.assertTrue(self.generator.validate_address("30001_10", "STR20"))
        self.assertTrue(self.generator.validate_address("30001_10", "STRING"))

        # BITS
        self.assertTrue(self.generator.validate_address("30001_0_1", "BITS"))

        # Range check still applies to base address of compound
        self.assertFalse(self.generator.validate_address("65536_10", "STR20"))

    def test_validate_csv(self):
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.csv') as tmp:
            writer = csv.writer(tmp, delimiter=';')
            writer.writerow(['protocol', 'category', 'mfg', 'model', '', '', '', '', '', '', ''])
            # Valid row
            writer.writerow(['1', '3', '40001', 'U16', '', 'Name', 'tag1', '1.0', '0.0', 'Unit', '4'])
            # Duplicate tag
            writer.writerow(['2', '3', '40002', 'U16', '', 'Name2', 'tag1', '1.0', '0.0', 'Unit', '4'])
            tmp_path = tmp.name

        try:
            # Should fail due to duplicate tag
            self.assertFalse(self.generator.validate_csv(tmp_path))
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def test_validate_csv_overlaps(self):
         with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.csv') as tmp:
            writer = csv.writer(tmp, delimiter=';')
            writer.writerow(['protocol', 'category', 'mfg', 'model', '', '', '', '', '', '', ''])
            # Row 1: Address 100, Type U32 (uses 100, 101)
            writer.writerow(['1', '3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'W', '4'])
            # Row 2: Address 101, Type U16 (overlaps with 101)
            writer.writerow(['2', '3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'W', '4'])
            tmp_path = tmp.name

         try:
            # Overlaps are warnings, validate_csv currently returns True for overlaps
            # (unless I change it to return False). Let's check current implementation.
            # My implementation: self._check_address_overlap(...) is called but doesn't set valid = False.
            self.assertTrue(self.generator.validate_csv(tmp_path))
         finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def test_validate_csv_invalid_address(self):
         with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.csv') as tmp:
            writer = csv.writer(tmp, delimiter=';')
            writer.writerow(['protocol', 'category', 'mfg', 'model', '', '', '', '', '', '', ''])
            # Invalid address (out of range)
            writer.writerow(['1', '3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'W', '4'])
            tmp_path = tmp.name

         try:
            self.assertFalse(self.generator.validate_csv(tmp_path))
         finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

if __name__ == '__main__':
    unittest.main()
