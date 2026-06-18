import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_def = "test_validate.csv"

    def tearDown(self):
        if os.path.exists(self.test_def):
            os.remove(self.test_def)

    def test_validate_address_range(self):
        # Valid
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        self.assertTrue(self.generator.validate_address("0x100", "U16"))
        self.assertTrue(self.generator.validate_address("100_20", "STR20"))
        self.assertTrue(self.generator.validate_address("100_0_1", "BITS"))

        # Invalid range
        self.assertFalse(self.generator.validate_address("65536", "U16"))
        self.assertFalse(self.generator.validate_address("-1", "U16"))
        self.assertFalse(self.generator.validate_address("0x10000", "U16"))

    def test_validate_csv_duplicates(self):
        with open(self.test_def, 'w', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U16', '', 'Var1', 'tag1', '1', '0', 'V', '4'])
            writer.writerow(['2', '3', '101', 'U16', '', 'Var2', 'tag1', '1', '0', 'V', '4']) # Duplicate Tag

        self.assertFalse(self.generator.validate_csv(self.test_def))

    def test_validate_csv_overlaps(self):
        with open(self.test_def, 'w', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U32', '', 'Var1', 'tag1', '1', '0', 'V', '4']) # 100, 101
            writer.writerow(['2', '3', '101', 'U16', '', 'Var2', 'tag2', '1', '0', 'V', '4']) # Overlaps 101

        # Overlaps are currently warnings in _check_address_overlap, so validate_csv might still return True if it only checks fatal errors.
        # Looking at my implementation, it only sets valid=False for validate_address and duplicate tags.
        # Overlaps just log warnings.
        self.assertTrue(self.generator.validate_csv(self.test_def))

    def test_validate_csv_invalid_addr(self):
        with open(self.test_def, 'w', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '70000', 'U16', '', 'Var1', 'tag1', '1', '0', 'V', '4'])

        self.assertFalse(self.generator.validate_csv(self.test_def))

if __name__ == "__main__":
    unittest.main()
