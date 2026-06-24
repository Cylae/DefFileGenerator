import unittest
import os
import csv
import tempfile
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_dir = tempfile.TemporaryDirectory()
        # Suppress logging during tests
        logging.getLogger().setLevel(logging.CRITICAL)

    def tearDown(self):
        self.test_dir.cleanup()

    def test_validate_address_range(self):
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        self.assertFalse(self.generator.validate_address("65536", "U16"))
        self.assertFalse(self.generator.validate_address("-1", "U16"))

        # Hex addresses
        self.assertTrue(self.generator.validate_address("0xFFFF", "U16"))
        self.assertFalse(self.generator.validate_address("0x10000", "U16"))

        # Compound addresses
        self.assertTrue(self.generator.validate_address("100_20", "STRING"))
        self.assertFalse(self.generator.validate_address("70000_20", "STRING"))

    def test_validate_csv_duplicates(self):
        path = os.path.join(self.test_dir.name, "dup.csv")
        with open(path, 'w', newline='', encoding='utf-8') as f:
            f.write("modbusRTU;Inverter;MFG;MODEL;;;;;;;\n")
            f.write("1;3;100;U16;;Var1;tag1;1.0;0.0;V;4\n")
            f.write("2;3;101;U16;;Var2;tag1;1.0;0.0;V;4\n")

        self.assertFalse(self.generator.validate_csv(path))

    def test_validate_csv_valid(self):
        path = os.path.join(self.test_dir.name, "valid.csv")
        with open(path, 'w', newline='', encoding='utf-8') as f:
            f.write("modbusRTU;Inverter;MFG;MODEL;;;;;;;\n")
            f.write("1;3;100;U16;;Var1;tag1;1.0;0.0;V;4\n")
            f.write("2;3;101;U16;;Var2;tag2;1.0;0.0;V;4\n")

        self.assertTrue(self.generator.validate_csv(path))

    def test_validate_csv_overlap(self):
        path = os.path.join(self.test_dir.name, "overlap.csv")
        with open(path, 'w', newline='', encoding='utf-8') as f:
            f.write("modbusRTU;Inverter;MFG;MODEL;;;;;;;\n")
            f.write("1;3;100;U32;;Var1;tag1;1.0;0.0;V;4\n")
            f.write("2;3;101;U16;;Var2;tag2;1.0;0.0;V;4\n")

        # Overlaps are logged as warnings but validate_csv currently returns True for overlaps
        # based on my implementation (it only sets valid=False for Fatal or Invalid Type/Addr).
        # Let's check my implementation again.
        # It calls _check_address_overlap which logs warnings.
        self.assertTrue(self.generator.validate_csv(path))

if __name__ == '__main__':
    unittest.main()
