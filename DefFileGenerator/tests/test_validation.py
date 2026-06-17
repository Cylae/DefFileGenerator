import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_file = "test_validate.csv"
        logging.basicConfig(level=logging.ERROR)

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def create_def_file(self, rows):
        with open(self.test_file, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)

    def test_valid_file(self):
        rows = [
            ['3', '30001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '30002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']
        ]
        self.create_def_file(rows)
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_duplicate_tag(self):
        rows = [
            ['3', '30001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '30002', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'A', '4']
        ]
        self.create_def_file(rows)
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_invalid_address_range(self):
        rows = [
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.create_def_file(rows)
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_address_overlap_warning(self):
        # Overlaps currently return True but log warnings in Generator._check_address_overlap
        # unless we change validate_csv to be stricter. The prompt says "duplicate tag is fatal".
        rows = [
            ['3', '30001', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '30002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']
        ]
        self.create_def_file(rows)
        # Should still be True as per current implementation (overlaps are warnings)
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_address_range_check(self):
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        self.assertFalse(self.generator.validate_address("65536", "U16"))
        self.assertFalse(self.generator.validate_address("-1", "U16"))
        self.assertTrue(self.generator.validate_address("0x0", "U16"))
        self.assertTrue(self.generator.validate_address("0xFFFF", "U16"))

if __name__ == '__main__':
    unittest.main()
