import unittest
import os
import tempfile
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.temp_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.temp_dir.cleanup()

    def create_csv(self, rows, filename="test.csv"):
        filepath = os.path.join(self.temp_dir.name, filename)
        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)
        return filepath

    def test_validate_success(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        filepath = self.create_csv(rows)
        self.assertTrue(self.generator.validate_csv(filepath))

    def test_duplicate_tag(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        filepath = self.create_csv(rows)
        with self.assertLogs(level='ERROR') as cm:
            self.assertFalse(self.generator.validate_csv(filepath))
            self.assertTrue(any("Duplicate Tag 'tag1'" in msg for msg in cm.output))

    def test_address_overlap(self):
        rows = [
            ['3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ]
        filepath = self.create_csv(rows)
        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_csv(filepath))
            self.assertTrue(any("Address overlap detected" in msg for msg in cm.output))

    def test_address_out_of_range(self):
        rows = [
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        filepath = self.create_csv(rows)
        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_csv(filepath))
            self.assertTrue(any("Address 70000 is out of standard Modbus range" in msg for msg in cm.output))

if __name__ == '__main__':
    unittest.main()
