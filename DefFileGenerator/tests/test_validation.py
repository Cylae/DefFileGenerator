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
        logging.basicConfig(level=logging.ERROR)

    def tearDown(self):
        self.test_dir.cleanup()

    def create_def_file(self, rows):
        path = os.path.join(self.test_dir.name, 'test_def.csv')
        with open(path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + row)
        return path

    def test_validate_address_range(self):
        # Standard range 0-65535
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        self.assertTrue(self.generator.validate_address("0x0", "U16"))
        self.assertTrue(self.generator.validate_address("0xFFFF", "U16"))

        # Out of range
        self.assertFalse(self.generator.validate_address("-1", "U16"))
        self.assertFalse(self.generator.validate_address("65536", "U16"))
        self.assertFalse(self.generator.validate_address("0x10000", "U16"))

    def test_validate_str_synonym(self):
        # STR20 is a synonym for STRING
        self.assertTrue(self.generator.validate_address("100_20", "STR20"))
        self.assertTrue(self.generator.validate_address("100_20", "STRING"))
        self.assertFalse(self.generator.validate_address("100", "STR20"))

    def test_validate_csv_success(self):
        # Info1, Info2, Info3, Info4, Name, Tag, CoefA, CoefB, Unit, Action
        rows = [
            ['3', '100', 'U16', '', 'Var1', 'var1', '1.0', '0.0', 'V', '1'],
            ['4', '200', 'F32', '', 'Var2', 'var2', '1.0', '0.0', 'A', '4']
        ]
        path = self.create_def_file(rows)
        self.assertTrue(self.generator.validate_csv(path))

    def test_validate_csv_overlap(self):
        # Overlapping registers
        rows = [
            ['3', '100', 'U32', '', 'Var1', 'var1', '1.0', '0.0', 'V', '1'],
            ['3', '101', 'U16', '', 'Var2', 'var2', '1.0', '0.0', 'A', '4']
        ]
        path = self.create_def_file(rows)
        self.assertFalse(self.generator.validate_csv(path))

    def test_validate_csv_invalid_address(self):
        rows = [
            ['3', '70000', 'U16', '', 'Var1', 'var1', '1.0', '0.0', 'V', '1']
        ]
        path = self.create_def_file(rows)
        self.assertFalse(self.generator.validate_csv(path))

    def test_validate_csv_bits_no_overlap(self):
        # Multiple BITS on same base address is allowed
        rows = [
            ['3', '100_0_1', 'BITS', '', 'Bit1', 'bit1', '1.0', '0.0', '', '1'],
            ['3', '100_1_1', 'BITS', '', 'Bit2', 'bit2', '1.0', '0.0', '', '1']
        ]
        path = self.create_def_file(rows)
        self.assertTrue(self.generator.validate_csv(path))

if __name__ == '__main__':
    unittest.main()
