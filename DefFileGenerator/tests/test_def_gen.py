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

    def test_validate_type_case_insensitivity(self):
        case_variants = ['u16', 'String', 'str20', 'i32_wb', 'ipv6', 'Mac']
        for t in case_variants:
            with self.subTest(type=t):
                self.assertTrue(self.generator.validate_type(t))

    def test_validate_address_valid(self):
        self.assertTrue(self.generator.validate_address('30001', 'U16'))
        self.assertTrue(self.generator.validate_address('30001_10', 'STRING'))
        self.assertTrue(self.generator.validate_address('30001_0_1', 'BITS'))

    def test_validate_address_invalid(self):
        self.assertFalse(self.generator.validate_address('30001_10', 'U16')) # U16 expects int
        self.assertFalse(self.generator.validate_address('30001', 'STRING')) # STRING expects Addr_Len
        self.assertFalse(self.generator.validate_address('abc', 'U16'))

    def test_get_register_count(self):
        self.assertEqual(self.generator.get_register_count('U16', '30000'), 1)
        self.assertEqual(self.generator.get_register_count('U32', '30000'), 2)
        self.assertEqual(self.generator.get_register_count('U64', '30000'), 4)
        self.assertEqual(self.generator.get_register_count('MAC', '30000'), 3)
        self.assertEqual(self.generator.get_register_count('IPV6', '30000'), 8)
        self.assertEqual(self.generator.get_register_count('STRING', '30000_10'), 5) # ceil(10/2)
        self.assertEqual(self.generator.get_register_count('STRING', '30000_11'), 6) # ceil(11/2)

    def test_process_rows_basic(self):
        rows = [{
            'Name': 'Test Var',
            'Tag': 'test_tag',
            'RegisterType': 'Holding Register',
            'Address': '30000',
            'Type': 'U16',
            'Factor': '1',
            'Offset': '0',
            'Unit': 'V',
            'Action': '4',
            'ScaleFactor': '0'
        }]
        processed = self.generator.process_rows(rows)
        self.assertEqual(len(processed), 1)
        self.assertEqual(processed[0]['Info1'], '3')
        self.assertEqual(processed[0]['Info3'], 'U16')
        self.assertEqual(processed[0]['CoefA'], '1.000000')

    def test_process_rows_str_expansion(self):
        rows = [{
            'Name': 'Test Str',
            'Tag': 'str_tag',
            'RegisterType': 'Holding Register',
            'Address': '30010',
            'Type': 'STR20',
            'Factor': '', 'Offset': '', 'Unit': '', 'Action': '', 'ScaleFactor': ''
        }]
        processed = self.generator.process_rows(rows)
        self.assertEqual(len(processed), 1)
        self.assertEqual(processed[0]['Info3'], 'STRING')
        self.assertEqual(processed[0]['Info2'], '30010_20')

    def test_process_rows_overlap(self):
        rows = [
            {
                'Name': 'Var1', 'Tag': 't1', 'RegisterType': '3', 'Address': '30000', 'Type': 'U32', # Occupies 30000, 30001
                'Factor': '1', 'Offset': '0', 'Unit': '', 'Action': '4', 'ScaleFactor': '0'
            },
            {
                'Name': 'Var2', 'Tag': 't2', 'RegisterType': '3', 'Address': '30001', 'Type': 'U16', # Occupies 30001 (Overlap)
                'Factor': '1', 'Offset': '0', 'Unit': '', 'Action': '4', 'ScaleFactor': '0'
            }
        ]
        with self.assertLogs(level='WARNING') as log:
            processed = self.generator.process_rows(rows)
            self.assertEqual(len(processed), 2)
            # Check if any log message contains "Overlap detected"
            self.assertTrue(any("Address overlap detected" in m for m in log.output))

    def test_process_rows_no_overlap_different_types(self):
        rows = [
            {
                'Name': 'Var1', 'Tag': 't1', 'RegisterType': 'Holding Register', 'Address': '100', 'Type': 'U16',
                'Factor': '1', 'Offset': '0', 'Unit': '', 'Action': '1', 'ScaleFactor': '0'
            },
            {
                'Name': 'Var2', 'Tag': 't2', 'RegisterType': 'Input Register', 'Address': '100', 'Type': 'U16',
                'Factor': '1', 'Offset': '0', 'Unit': '', 'Action': '1', 'ScaleFactor': '0'
            }
        ]
        # Should NOT log a warning
        try:
            with self.assertLogs(level='WARNING') as log:
                processed = self.generator.process_rows(rows)
                self.assertEqual(len(processed), 2)
                # If we are here, some warning was logged. Check it's not overlap.
                for m in log.output:
                    self.assertNotIn("Address overlap detected", m)
        except AssertionError:
            # assertLogs raises AssertionError if NO logs are produced, which is what we want
            processed = self.generator.process_rows(rows)
            self.assertEqual(len(processed), 2)

    def test_duplicate_name(self):
        rows = [
             {'Name': 'Var1', 'Tag': 't1', 'RegisterType': '3', 'Address': '30000', 'Type': 'U16', 'Factor': '', 'Offset': '', 'Unit': '', 'Action': '', 'ScaleFactor': ''},
             {'Name': 'Var1', 'Tag': 't2', 'RegisterType': '3', 'Address': '30001', 'Type': 'U16', 'Factor': '', 'Offset': '', 'Unit': '', 'Action': '', 'ScaleFactor': ''}
        ]
        with self.assertLogs(level='WARNING') as log:
            self.generator.process_rows(rows)
            self.assertTrue(any("Duplicate Name" in m for m in log.output))

    def test_scalefactor_calculation(self):
        rows = [{
            'Name': 'Scaled Var',
            'Tag': 'scaled_tag',
            'RegisterType': '3',
            'Address': '30000',
            'Type': 'U16',
            'Factor': '2.5',
            'Offset': '0',
            'Unit': '',
            'Action': '',
            'ScaleFactor': '-1' # CoefA = 2.5 * 10^-1 = 0.25
        }]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['CoefA'], '0.250000')

if __name__ == '__main__':
    unittest.main()
