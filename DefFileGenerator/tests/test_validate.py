import unittest
import logging
import os
import csv
import tempfile
from DefFileGenerator.def_gen import Generator, peek_generator

class TestGeneratorValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.test_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.test_dir.cleanup()

    def test_peek_generator(self):
        # Empty
        has_data, it = peek_generator([])
        self.assertFalse(has_data)
        self.assertEqual(list(it), [])

        # Non-empty
        has_data, it = peek_generator([1, 2, 3])
        self.assertTrue(has_data)
        self.assertEqual(list(it), [1, 2, 3])

        # None
        has_data, it = peek_generator(None)
        self.assertFalse(has_data)
        self.assertEqual(list(it), [])

    def test_validate_address_range(self):
        # Valid
        self.assertTrue(self.generator.validate_address('0', 'U16'))
        self.assertTrue(self.generator.validate_address('65535', 'U16'))
        self.assertTrue(self.generator.validate_address('0x0', 'U16'))
        self.assertTrue(self.generator.validate_address('0xFFFF', 'U16'))

        # Out of range (should log warning but return True for format validity)
        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(self.generator.validate_address('65536', 'U16'))
            self.assertTrue(any("outside standard Modbus range" in m for m in log.output))

        with self.assertLogs(level='WARNING') as log:
            self.assertTrue(self.generator.validate_address('-1', 'U16'))
            self.assertTrue(any("outside standard Modbus range" in m for m in log.output))

    def test_validate_csv_full(self):
        path = os.path.join(self.test_dir.name, 'test_def.csv')
        with open(path, 'w', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'Model', '', '', '', '', '', '', ''])
            # 1: Valid
            writer.writerow(['1', '3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])
            # 2: Invalid Type
            writer.writerow(['2', '3', '101', 'INVALID', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4'])
            # 3: Invalid Address
            writer.writerow(['3', '3', 'not_an_addr', 'U16', '', 'Var3', 'tag3', '1.0', '0.0', 'V', '4'])

        with self.assertLogs(level='WARNING') as log:
            valid = self.generator.validate_csv(path)
            self.assertFalse(valid)
            self.assertTrue(any("Invalid Type 'INVALID'" in m for m in log.output))
            self.assertTrue(any("Invalid Address 'not_an_addr'" in m for m in log.output))

    def test_validate_csv_duplicate_tag(self):
        path = os.path.join(self.test_dir.name, 'test_dup.csv')
        with open(path, 'w', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'Model', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])
            writer.writerow(['2', '3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4'])

        with self.assertLogs(level='ERROR') as log:
            valid = self.generator.validate_csv(path)
            self.assertFalse(valid)
            self.assertTrue(any("Duplicate Tag 'tag1' detected" in m for m in log.output))

if __name__ == '__main__':
    unittest.main()
