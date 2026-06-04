import unittest
import logging
import os
import csv
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.gen = Generator()
        self.test_file = "test_validate.csv"

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def test_validate_success(self):
        with open(self.test_file, 'w', encoding='utf-8') as f:
            f.write("modbusRTU;Inverter;MFG;MODEL;;;;;;;\n")
            f.write("1;3;100;U16;;Name1;tag1;1.0;0.0;V;4\n")
            f.write("2;3;101;U16;;Name2;tag2;1.0;0.0;V;4\n")

        self.assertTrue(self.gen.validate_csv(self.test_file))

    def test_validate_duplicate_tag(self):
        with open(self.test_file, 'w', encoding='utf-8') as f:
            f.write("modbusRTU;Inverter;MFG;MODEL;;;;;;;\n")
            f.write("1;3;100;U16;;Name1;tag1;1.0;0.0;V;4\n")
            f.write("2;3;101;U16;;Name2;tag1;1.0;0.0;V;4\n")

        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(self.gen.validate_csv(self.test_file))
            self.assertTrue(any("Duplicate Tag 'tag1'" in m for m in log.output))

    def test_validate_overlap(self):
        with open(self.test_file, 'w', encoding='utf-8') as f:
            f.write("modbusRTU;Inverter;MFG;MODEL;;;;;;;\n")
            f.write("1;3;100;U32;;Name1;tag1;1.0;0.0;V;4\n")
            f.write("2;3;101;U16;;Name2;tag2;1.0;0.0;V;4\n")

        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(self.gen.validate_csv(self.test_file))
            self.assertTrue(any("Address overlap detected" in m for m in log.output))

if __name__ == "__main__":
    unittest.main()
