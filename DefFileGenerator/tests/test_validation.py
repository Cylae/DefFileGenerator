import unittest
import os
import csv
import tempfile
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.temp_dir = tempfile.TemporaryDirectory()
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        self.temp_dir.cleanup()
        logging.disable(logging.NOTSET)

    def test_validate_address_range(self):
        # Valid addresses
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        self.assertTrue(self.generator.validate_address("0x100", "U16"))
        self.assertTrue(self.generator.validate_address("100h", "U16"))

        # Invalid addresses (out of range)
        self.assertFalse(self.generator.validate_address("-1", "U16"))
        self.assertFalse(self.generator.validate_address("65536", "U16"))
        self.assertFalse(self.generator.validate_address("0x10000", "U16"))

    def test_validate_address_str_n(self):
        # STR<n> should be validated as STRING
        self.assertTrue(self.generator.validate_address("30030_20", "STR20"))
        self.assertFalse(self.generator.validate_address("30030", "STR20")) # Missing length in address

    def test_validate_csv_format(self):
        # Create a valid definition file
        valid_file = os.path.join(self.temp_dir.name, "valid.csv")
        with open(valid_file, 'w', encoding='utf-8-sig') as f:
            f.write("modbusRTU;Inverter;MFG;MODEL;;;;;;;\n")
            f.write("1;3;100;U16;;Test;test;1.0;0.0;V;4\n")

        self.assertTrue(self.generator.validate_csv(valid_file))

        # Create an invalid definition file (invalid address)
        invalid_file = os.path.join(self.temp_dir.name, "invalid.csv")
        with open(invalid_file, 'w', encoding='utf-8-sig') as f:
            f.write("modbusRTU;Inverter;MFG;MODEL;;;;;;;\n")
            f.write("1;3;70000;U16;;Test;test;1.0;0.0;V;4\n")

        self.assertFalse(self.generator.validate_csv(invalid_file))

    def test_validate_csv_missing_columns(self):
        # Row with missing columns
        missing_cols_file = os.path.join(self.temp_dir.name, "missing_cols.csv")
        with open(missing_cols_file, 'w', encoding='utf-8-sig') as f:
            f.write("modbusRTU;Inverter;MFG;MODEL;;;;;;;\n")
            f.write("1;3;100;U16\n")

        self.assertFalse(self.generator.validate_csv(missing_cols_file))

if __name__ == '__main__':
    unittest.main()
