import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_file = "test_validation_def.csv"
        # Suppress logging during tests unless needed
        logging.getLogger().setLevel(logging.ERROR)

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def write_csv(self, rows):
        with open(self.test_file, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + list(row))

    def test_valid_file(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']
        ]
        self.write_csv(rows)
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_invalid_type(self):
        rows = [
            ['3', '100', 'INVALID_TYPE', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.write_csv(rows)
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_invalid_address_format(self):
        rows = [
            ['3', 'not_a_number', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.write_csv(rows)
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_address_out_of_range(self):
        rows = [
            ['3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        self.write_csv(rows)
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_duplicate_tag(self):
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'tag_dup', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag_dup', '1.0', '0.0', 'A', '4']
        ]
        self.write_csv(rows)
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_address_overlap(self):
        rows = [
            ['3', '100', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['3', '101', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']
        ]
        self.write_csv(rows)
        self.assertFalse(self.generator.validate_csv(self.test_file))

    def test_bits_no_overlap_same_address(self):
        rows = [
            ['3', '100_0_1', 'BITS', '', 'Bit1', 'tag1', '1.0', '0.0', '', '4'],
            ['3', '100_1_1', 'BITS', '', 'Bit2', 'tag2', '1.0', '0.0', '', '4']
        ]
        self.write_csv(rows)
        self.assertTrue(self.generator.validate_csv(self.test_file))

    def test_str_n_synonym_validation(self):
        # validate_address should handle STR<n>
        self.assertTrue(self.generator.validate_address("3000_10", "STR10"))
        self.assertFalse(self.generator.validate_address("3000", "STR10"))

if __name__ == '__main__':
    unittest.main()
