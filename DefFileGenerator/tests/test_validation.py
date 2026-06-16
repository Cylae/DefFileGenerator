import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_file = 'test_validation.csv'

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def test_validate_csv_valid(self):
        with open(self.test_file, 'w', encoding='utf-8') as f:
            f.write("modbusRTU;Inverter;Test;Model;;;;;;;\n")
            f.write("1;3;40001;U16;;Name1;tag1;1.0;0.0;V;4\n")
            f.write("2;4;30001;I32;;Name2;tag2;1.0;0.0;A;4\n")

        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_validate_csv_invalid_tag(self):
        with open(self.test_file, 'w', encoding='utf-8') as f:
            f.write("modbusRTU;Inverter;Test;Model;;;;;;;\n")
            f.write("1;3;40001;U16;;Name1;tag1;1.0;0.0;V;4\n")
            f.write("2;3;40002;U16;;Name2;tag1;1.0;0.0;V;4\n") # Duplicate tag1

        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("Duplicate Tag 'tag1'" in m for m in log.output))

    def test_validate_csv_invalid_address(self):
        with open(self.test_file, 'w', encoding='utf-8') as f:
            f.write("modbusRTU;Inverter;Test;Model;;;;;;;\n")
            f.write("1;3;70000;U16;;Name1;tag1;1.0;0.0;V;4\n") # Address out of range

        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(self.generator.validate_csv(self.test_file))
            self.assertTrue(any("out of Modbus range" in m for m in log.output))

    def test_action_defaulting(self):
        rows = [
            {'Name': 'Holding', 'Address': '40001', 'RegisterType': 'Holding Register', 'Type': 'U16'},
            {'Name': 'Input', 'Address': '30001', 'RegisterType': 'Input Register', 'Type': 'U16'},
            {'Name': 'Coil', 'Address': '1', 'RegisterType': 'Coil', 'Type': 'U16'},
            {'Name': 'Discrete', 'Address': '1', 'RegisterType': 'Discrete Input', 'Type': 'U16'},
        ]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1') # Holding -> 1
        self.assertEqual(processed[1]['Action'], '4') # Input -> 4
        self.assertEqual(processed[2]['Action'], '1') # Coil -> 1
        self.assertEqual(processed[3]['Action'], '4') # Discrete -> 4

if __name__ == '__main__':
    unittest.main()
