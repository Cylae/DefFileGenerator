import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestGenerator(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_validate_type_valid(self):
        valid_types = [
            'U16', 'I32', 'F32', 'STRING', 'BITS', 'IP', 'MAC', 'STR10', 'U16_W', 'I32_WB',
            'U8', 'I8', 'U64', 'I64', 'F64', 'IPV6', 'U16_B', 'U32_W', 'I64_WB'
        ]
        for t in valid_types:
            with self.subTest(type=t):
                self.assertTrue(self.generator.validate_type(t))

    def test_validate_type_invalid(self):
        invalid_types = [
            'UNKNOWN', 'U1', 'I128', 'STR', 'BITS_2',
            'F32_W', 'F64_WB', 'STRING_W', 'U17', 'I31'
        ]
        for t in invalid_types:
            with self.subTest(type=t):
                self.assertFalse(self.generator.validate_type(t))

    def test_normalize_type(self):
        self.assertEqual(self.generator.normalize_type("uint16"), "U16")
        self.assertEqual(self.generator.normalize_type("int32"), "I32")
        self.assertEqual(self.generator.normalize_type("float"), "F32")
        self.assertEqual(self.generator.normalize_type("float64"), "F64")
        self.assertEqual(self.generator.normalize_type("U16_W"), "U16_W")
        self.assertEqual(self.generator.normalize_type("unsigned int 16"), "U16")

    def test_normalize_action(self):
        self.assertEqual(self.generator.normalize_action("R"), "4")
        self.assertEqual(self.generator.normalize_action("RW"), "1")
        self.assertEqual(self.generator.normalize_action("Read"), "4")
        self.assertEqual(self.generator.normalize_action("Write"), "1")
        self.assertEqual(self.generator.normalize_action("7"), "7")
        self.assertEqual(self.generator.normalize_action(""), "1")

    def test_validate_address_valid(self):
        self.assertTrue(self.generator.validate_address('30001', 'U16'))
        self.assertTrue(self.generator.validate_address('-1', 'U16')) # Now allowed
        self.assertTrue(self.generator.validate_address('0x7531', 'U16'))
        self.assertTrue(self.generator.validate_address('7531h', 'U16'))
        self.assertTrue(self.generator.validate_address('A0', 'U16'))
        self.assertTrue(self.generator.validate_address('30001_10', 'STRING'))
        self.assertTrue(self.generator.validate_address('0x7531_10', 'STRING'))
        self.assertTrue(self.generator.validate_address('A0_10', 'STRING'))
        self.assertTrue(self.generator.validate_address('30001_0_1', 'BITS'))
        self.assertTrue(self.generator.validate_address('0x7531_0_1', 'BITS'))
        self.assertTrue(self.generator.validate_address('A0_0_1', 'BITS'))

    def test_normalize_address_val(self):
        self.assertEqual(self.generator.normalize_address_val('0x10'), '16')
        self.assertEqual(self.generator.normalize_address_val('10h'), '16')
        self.assertEqual(self.generator.normalize_address_val('10'), '10')
        self.assertEqual(self.generator.normalize_address_val('A0'), '160')
        self.assertEqual(self.generator.normalize_address_val('1,234'), '1234')
        self.assertEqual(self.generator.normalize_address_val('Reg: 0x10 (Holding)'), '16')

    def test_validate_address_invalid(self):
        self.assertFalse(self.generator.validate_address('30001_10', 'U16'))
        self.assertFalse(self.generator.validate_address('30001', 'STRING'))
        self.assertFalse(self.generator.validate_address('xyz', 'U16'))

    def test_get_register_count(self):
        self.assertEqual(self.generator.get_register_count('U16', '30000'), 1)
        self.assertEqual(self.generator.get_register_count('U32', '30000'), 2)
        self.assertEqual(self.generator.get_register_count('U64', '30000'), 4)
        self.assertEqual(self.generator.get_register_count('MAC', '30000'), 3)
        self.assertEqual(self.generator.get_register_count('IPV6', '30000'), 8)
        self.assertEqual(self.generator.get_register_count('STRING', '30000_10'), 5)
        self.assertEqual(self.generator.get_register_count('STRING', '30000_11'), 6)

    def test_process_rows_basic(self):
        rows = [{
            'Name': 'Test Var',
            'Tag': 'test_tag',
            'RegisterType': 'Holding Register',
            'Address': '30000',
            'Type': 'uint16', # Tests normalization
            'Factor': '1',
            'Offset': '0',
            'Unit': 'V',
            'Action': 'Read', # Tests normalization
            'ScaleFactor': '0'
        }]
        processed = self.generator.process_rows(rows)
        self.assertEqual(len(processed), 1)
        self.assertEqual(processed[0]['Info1'], '3')
        self.assertEqual(processed[0]['Info3'], 'U16')
        self.assertEqual(processed[0]['Action'], '4')
        self.assertEqual(processed[0]['CoefA'], '1.000000')

    def test_address_offset(self):
        self.generator.address_offset = 1
        rows = [{
            'Name': 'Test Var',
            'Address': '30001',
            'Type': 'U16'
        }]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '30000')

    def test_negative_address_warning(self):
        self.generator.address_offset = 100
        rows = [{
            'Name': 'Test Var',
            'Address': '50',
            'Type': 'U16'
        }]
        with self.assertLogs(level='WARNING') as log:
            processed = self.generator.process_rows(rows)
            self.assertEqual(processed[0]['Info2'], '-50')
            self.assertTrue(any("results in negative address -50" in m for m in log.output))

    def test_process_rows_overlap(self):
        rows = [
            {
                'Name': 'Var1', 'Tag': 't1', 'RegisterType': '3', 'Address': '30000', 'Type': 'U32',
                'Factor': '1', 'Offset': '0', 'Unit': '', 'Action': '4', 'ScaleFactor': '0'
            },
            {
                'Name': 'Var2', 'Tag': 't2', 'RegisterType': '3', 'Address': '30001', 'Type': 'U16',
                'Factor': '1', 'Offset': '0', 'Unit': '', 'Action': '4', 'ScaleFactor': '0'
            }
        ]
        with self.assertLogs(level='WARNING') as log:
            processed = self.generator.process_rows(rows)
            self.assertEqual(len(processed), 2)
            self.assertTrue(any("Address overlap detected" in m for m in log.output))

    def test_duplicate_name_warning_format(self):
        rows = [
             {'Name': 'Var1', 'Address': '30000', 'Type': 'U16'},
             {'Name': 'Var1', 'Address': '30001', 'Type': 'U16'}
        ]
        with self.assertLogs(level='WARNING') as log:
            self.generator.process_rows(rows)
            self.assertTrue(any("Duplicate Name 'Var1' detected. Previous occurrence at line 2." in m for m in log.output))

if __name__ == '__main__':
    unittest.main()
