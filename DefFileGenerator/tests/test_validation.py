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

    def create_csv(self, rows, delimiter=';'):
        path = os.path.join(self.test_dir.name, 'test.csv')
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=delimiter)
            for row in rows:
                writer.writerow(row)
        return path

    def test_validate_valid_csv(self):
        rows = [
            ['modbusRTU', 'Inverter', 'Mfg', 'Model', '', '', '', '', '', '', ''],
            ['1', '3', '40001', 'U16', '', 'Name1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_csv(rows)
        self.assertTrue(self.generator.validate_csv(path))

    def test_validate_invalid_type(self):
        rows = [
            ['modbusRTU', 'Inverter', 'Mfg', 'Model', '', '', '', '', '', '', ''],
            ['1', '3', '40001', 'INVALID_TYPE', '', 'Name1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_csv(rows)
        self.assertFalse(self.generator.validate_csv(path))

    def test_validate_out_of_range_address(self):
        rows = [
            ['modbusRTU', 'Inverter', 'Mfg', 'Model', '', '', '', '', '', '', ''],
            ['1', '3', '70000', 'U16', '', 'Name1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_csv(rows)
        self.assertFalse(self.generator.validate_csv(path))

    def test_validate_hex_address(self):
        rows = [
            ['modbusRTU', 'Inverter', 'Mfg', 'Model', '', '', '', '', '', '', ''],
            ['1', '3', '0x9C40', 'U16', '', 'Name1', 'tag1', '1.0', '0.0', 'V', '4']
        ]
        path = self.create_csv(rows)
        self.assertTrue(self.generator.validate_csv(path)) # 0x9C40 = 40000 (valid)

    def test_validate_invalid_row_length(self):
        rows = [
            ['modbusRTU', 'Inverter', 'Mfg', 'Model', '', '', '', '', '', '', ''],
            ['1', '3', '40001', 'U16']
        ]
        path = self.create_csv(rows)
        self.assertFalse(self.generator.validate_csv(path))

    def test_validate_address_range_logic(self):
        self.assertTrue(self.generator.validate_address('0', 'U16'))
        self.assertTrue(self.generator.validate_address('65535', 'U16'))
        self.assertFalse(self.generator.validate_address('65536', 'U16'))
        self.assertFalse(self.generator.validate_address('-1', 'U16'))

    def test_validate_compound_address(self):
        self.assertTrue(self.generator.validate_address('40001_20', 'STRING'))
        self.assertTrue(self.generator.validate_address('40001_1_1', 'BITS'))
        self.assertFalse(self.generator.validate_address('70000_20', 'STRING'))

    def test_validate_str_synonym(self):
        self.assertTrue(self.generator.validate_address('40001_20', 'STR20'))
        self.assertFalse(self.generator.validate_address('70000_20', 'STR20'))

if __name__ == '__main__':
    unittest.main()
