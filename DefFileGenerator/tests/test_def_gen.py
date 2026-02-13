import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestGenerator(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_validate_type_valid(self):
        valid_types = ['U16', 'I32', 'F32', 'STRING', 'BITS', 'IP', 'MAC', 'STR10', 'U16_W', 'I32_WB', 'U8', 'I8', 'U64', 'I64', 'F64', 'IPV6', 'U16_B', 'U32_W', 'I64_WB']
        for t in valid_types:
            with self.subTest(type=t):
                self.assertTrue(self.generator.validate_type(t))

    def test_validate_type_invalid(self):
        invalid_types = ['UNKNOWN', 'U1', 'I128', 'STR', 'BITS_2', 'F32_W', 'F64_WB', 'STRING_W', 'U17', 'I31']
        for t in invalid_types:
            with self.subTest(type=t):
                self.assertFalse(self.generator.validate_type(t))

    def test_normalize_type(self):
        cases = [
            ("uint16", "U16"), ("INT32", "I32"), ("float", "F32"), ("Double", "F64"),
            ("unsigned int 16", "U16"), ("signed long 32", "I32"), ("Uint32_W", "U32_W"),
            ("bool", "BITS"), ("boolean", "BITS"), ("str20", "STR20"),
            ("String (10 chars)", "STRING"), ("uint16 (holding)", "U16")
        ]
        for inp, expected in cases:
            with self.subTest(inp=inp):
                self.assertEqual(self.generator.normalize_type(inp), expected)

    def test_normalize_action(self):
        cases = [("R", "4"), ("Read", "4"), ("read-only", "4"), ("RW", "1"), ("write", "1"), ("Read/Write", "1"), ("1", "1"), ("4", "4"), ("unknown", "1")]
        for inp, expected in cases:
            with self.subTest(inp=inp):
                self.assertEqual(self.generator.normalize_action(inp), expected)

    def test_validate_address_valid(self):
        self.assertTrue(self.generator.validate_address('30001', 'U16'))
        self.assertTrue(self.generator.validate_address('0x7531', 'U16'))
        self.assertTrue(self.generator.validate_address('7531h', 'U16'))
        self.assertTrue(self.generator.validate_address('A0', 'U16'))
        self.assertTrue(self.generator.validate_address('30001_10', 'STRING'))
        self.assertTrue(self.generator.validate_address('30001_0_1', 'BITS'))

    def test_validate_address_invalid(self):
        self.assertFalse(self.generator.validate_address('30001_10', 'U16'))
        self.assertFalse(self.generator.validate_address('30001', 'STRING'))
        self.assertFalse(self.generator.validate_address('xyz', 'U16'))

    def test_normalize_address_val(self):
        self.assertEqual(self.generator.normalize_address_val('0x10'), '16')
        self.assertEqual(self.generator.normalize_address_val('10h'), '16')
        self.assertEqual(self.generator.normalize_address_val('10'), '10')
        self.assertEqual(self.generator.normalize_address_val('A0'), '160')
        self.assertEqual(self.generator.normalize_address_val('1,234'), '1234')

    def test_get_register_count(self):
        self.assertEqual(self.generator.get_register_count('U16', '30000'), 1)
        self.assertEqual(self.generator.get_register_count('U32', '30000'), 2)
        self.assertEqual(self.generator.get_register_count('U64', '30000'), 4)
        self.assertEqual(self.generator.get_register_count('STRING', '30000_10'), 5)

    def test_process_rows_robust(self):
        rows = [{'Name': 'Robust Var', 'RegisterType': 'holding', 'Address': '0x100', 'Type': 'uint16 (holding)', 'Action': 'read-only'}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(len(processed), 1)
        self.assertEqual(processed[0]['Info1'], '3')
        self.assertEqual(processed[0]['Info2'], '256')
        self.assertEqual(processed[0]['Action'], '4')

    def test_process_rows_str_expansion(self):
        rows = [{'Name': 'Test Str', 'Address': '30010', 'Type': 'STR20'}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Info3'], 'STRING')
        self.assertEqual(processed[0]['Info2'], '30010_20')

    def test_process_rows_overlap(self):
        rows = [{'Name': 'Var1', 'RegisterType': '3', 'Address': '30000', 'Type': 'U32'}, {'Name': 'Var2', 'RegisterType': '3', 'Address': '30001', 'Type': 'U16'}]
        with self.assertLogs(level='WARNING') as log:
            self.generator.process_rows(rows)
            self.assertTrue(any("Address overlap detected" in m for m in log.output))

    def test_process_rows_no_overlap_different_types(self):
        rows = [{'Name': 'Var1', 'RegisterType': '3', 'Address': '100', 'Type': 'U16'}, {'Name': 'Var2', 'RegisterType': '4', 'Address': '100', 'Type': 'U16'}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(len(processed), 2)

    def test_duplicate_detection(self):
        rows = [{'Name': 'Var1', 'Address': '30000', 'Type': 'U16'}, {'Name': 'Var1', 'Address': '30001', 'Type': 'U16'}]
        with self.assertLogs(level='WARNING') as log:
            self.generator.process_rows(rows)
            self.assertTrue(any("Duplicate Name" in m for m in log.output))

    def test_scalefactor_calculation(self):
        rows = [{'Name': 'S', 'Address': '1', 'Type': 'U16', 'Factor': '2.5', 'ScaleFactor': '-1'}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['CoefA'], '0.250000')

    def test_automatic_tag_generation(self):
        rows = [{'Name': 'T V', 'Address': '1', 'Type': 'U16'}, {'Name': 'T V', 'Address': '2', 'Type': 'U16'}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Tag'], 't_v')
        self.assertEqual(processed[1]['Tag'], 't_v_1')

if __name__ == '__main__':
    unittest.main()
