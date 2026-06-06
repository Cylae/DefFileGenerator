import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_file = 'test_validate_def.csv'
        logging.basicConfig(level=logging.INFO)

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def create_def_file(self, rows):
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';', lineterminator='\n')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)

    def test_valid_file(self):
        rows = [
            ['3', '30001', 'U16', '', 'Var1', 'tag1', '1.000000', '0.000000', 'V', '4'],
            ['3', '30002', 'U16', '', 'Var2', 'tag2', '1.000000', '0.000000', 'A', '4']
        ]
        self.create_def_file(rows)
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_duplicate_tag(self):
        rows = [
            ['3', '30001', 'U16', '', 'Var1', 'tag1', '1.000000', '0.000000', 'V', '4'],
            ['3', '30002', 'U16', '', 'Var2', 'tag1', '1.000000', '0.000000', 'A', '4']
        ]
        self.create_def_file(rows)
        with self.assertLogs(level='ERROR') as cm:
            result = self.generator.validate_csv(self.test_file)
            self.assertFalse(result)
            self.assertTrue(any("Duplicate Tag 'tag1'" in msg for msg in cm.output))

    def test_address_overlap(self):
        rows = [
            ['3', '30001', 'U32', '', 'Var1', 'tag1', '1.000000', '0.000000', 'V', '4'],
            ['3', '30002', 'U16', '', 'Var2', 'tag2', '1.000000', '0.000000', 'A', '4']
        ]
        self.create_def_file(rows)
        with self.assertLogs(level='WARNING') as cm:
            result = self.generator.validate_csv(self.test_file)
            self.assertTrue(result) # Overlap is a warning in this tool's validation logic
            self.assertTrue(any("Address overlap detected" in msg for msg in cm.output))

    def test_invalid_address_range(self):
        rows = [
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.000000', '0.000000', 'V', '4']
        ]
        self.create_def_file(rows)
        with self.assertLogs(level='WARNING') as cm:
            result = self.generator.validate_csv(self.test_file)
            self.assertTrue(result)
            self.assertTrue(any("out of standard Modbus range" in msg for msg in cm.output))

if __name__ == '__main__':
    unittest.main()
