import unittest
import os
import csv
import tempfile
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.test_dir.cleanup()

    def create_csv(self, rows):
        path = os.path.join(self.test_dir.name, "test.csv")
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)
        return path

    def test_valid_csv(self):
        path = self.create_csv([
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ])
        self.assertTrue(self.generator.validate_csv(path))

    def test_duplicate_tag(self):
        path = self.create_csv([
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        # Duplicate tags are FATAL (False)
        self.assertFalse(self.generator.validate_csv(path))

    def test_address_overlap(self):
        path = self.create_csv([
            ['3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4']
        ])
        # Overlaps are warnings (True, but logs warning)
        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_csv(path))
            self.assertTrue(any("overlap" in msg.lower() for msg in cm.output))

    def test_out_of_range_address(self):
        path = self.create_csv([
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_csv(path))
            self.assertTrue(any("outside standard range" in msg.lower() for msg in cm.output))

if __name__ == '__main__':
    unittest.main()
