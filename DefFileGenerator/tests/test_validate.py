import unittest
import os
import csv
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.filename = "test_validation.csv"

    def tearDown(self):
        if os.path.exists(self.filename):
            os.remove(self.filename)

    def test_validate_csv_success(self):
        with open(self.filename, 'w', encoding='utf-8') as f:
            f.write("modbusRTU;Inverter;Mfg;Model;;;;;;;\n")
            f.write("1;3;100;U16;;Name;tag;1.0;0.0;V;4\n")

        self.assertTrue(Generator.validate_csv(self.filename))

    def test_validate_csv_duplicate_tag(self):
        with open(self.filename, 'w', encoding='utf-8') as f:
            f.write("modbusRTU;Inverter;Mfg;Model;;;;;;;\n")
            f.write("1;3;100;U16;;Name1;tag;1.0;0.0;V;4\n")
            f.write("2;3;101;U16;;Name2;tag;1.0;0.0;V;4\n")

        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(Generator.validate_csv(self.filename))
            self.assertTrue(any("Duplicate Tag" in m for m in log.output))

    def test_validate_csv_address_overlap(self):
        with open(self.filename, 'w', encoding='utf-8') as f:
            f.write("modbusRTU;Inverter;Mfg;Model;;;;;;;\n")
            f.write("1;3;100;U32;;Name1;tag1;1.0;0.0;V;4\n")
            f.write("2;3;101;U16;;Name2;tag2;1.0;0.0;V;4\n")

        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(Generator.validate_csv(self.filename)) # Still valid, just warnings
            self.assertTrue(any("overlap detected" in m.lower() for m in log.output))

if __name__ == '__main__':
    unittest.main()
