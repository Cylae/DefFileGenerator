import unittest
import os
import csv
import logging
import tempfile
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_dir = tempfile.TemporaryDirectory()
        logging.basicConfig(level=logging.INFO)

    def tearDown(self):
        self.test_dir.cleanup()

    def create_csv(self, rows):
        path = os.path.join(self.test_dir.name, "test.csv")
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, 1):
                writer.writerow([str(i)] + row)
        return path

    def test_valid_csv(self):
        path = self.create_csv([
            ['3', '40001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '40002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ])
        self.assertTrue(self.generator.validate_csv(path))

    def test_duplicate_tag_fatal(self):
        path = self.create_csv([
            ['3', '40001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '40002', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        with self.assertLogs(level='ERROR') as log:
            self.assertFalse(self.generator.validate_csv(path))
            self.assertTrue(any("Duplicate Tag 'tag1' (Fatal)" in m for m in log.output))

    def test_address_overlap_warning(self):
        path = self.create_csv([
            ['3', '40001', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '40002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ])
        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(self.generator.validate_csv(path))
            self.assertTrue(any("Address overlap detected" in m for m in log.output))

    def test_invalid_address_range(self):
        path = self.create_csv([
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(self.generator.validate_csv(path))
            self.assertTrue(any("out of standard Modbus range" in m for m in log.output))

    def test_bom_handling(self):
        path = os.path.join(self.test_dir.name, "bom.csv")
        with open(path, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '40001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])
        self.assertTrue(self.generator.validate_csv(path))

if __name__ == '__main__':
    unittest.main()
