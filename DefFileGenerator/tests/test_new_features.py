import unittest
from DefFileGenerator.def_gen import Generator

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_intelligent_action_defaulting(self):
        # Info1 '2' (Discrete Input) and '4' (Input Register) should default to '4' (Read Only)
        # Info1 '1' (Coil) and '3' (Holding Register) should default to '1' (Read/Write)
        test_cases = [
            ('Discrete Input', '4'),
            ('Input Register', '4'),
            ('Coil', '1'),
            ('Holding Register', '1'),
            ('', '1'), # Defaults to Holding Register -> 1
        ]

        for reg_type, expected_action in test_cases:
            rows = [{
                'Name': 'TestVar',
                'Address': '100',
                'Type': 'U16',
                'RegisterType': reg_type,
                'Action': ''
            }]
            processed = list(self.generator.process_rows(rows))
            with self.subTest(reg_type=reg_type):
                self.assertEqual(processed[0]['Action'], expected_action)

    def test_address_range_validation(self):
        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_address("70000", "U16"))
            self.assertTrue(any("outside standard Modbus range" in output for output in cm.output))

    def test_bits_overlap_allowed(self):
        # BITS requires Address_BitOffset_Length format
        rows = [
            {'Name': 'Bit0', 'Address': '100_0_1', 'Type': 'BITS', 'RegisterType': 'Holding Register'},
            {'Name': 'Bit1', 'Address': '100_1_1', 'Type': 'BITS', 'RegisterType': 'Holding Register'}
        ]
        # Should not log overlap warning for BITS on same address
        with self.assertNoLogs(level='WARNING'):
            list(self.generator.process_rows(rows))

    def test_non_bits_overlap_warned(self):
        rows = [
            {'Name': 'Var1', 'Address': '100', 'Type': 'U16', 'RegisterType': 'Holding Register'},
            {'Name': 'Var2', 'Address': '100', 'Type': 'U16', 'RegisterType': 'Holding Register'}
        ]
        with self.assertLogs(level='WARNING') as cm:
            list(self.generator.process_rows(rows))
            self.assertTrue(any("overlap detected" in output for output in cm.output))

if __name__ == '__main__':
    unittest.main()
