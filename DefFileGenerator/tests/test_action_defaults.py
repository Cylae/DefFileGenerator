import unittest
from DefFileGenerator.def_gen import Generator

class TestActionDefaults(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_action_defaulting_based_on_type(self):
        test_cases = [
            # Holding Register (Info1: 3) -> Default to 1 (RW)
            ('Holding Register', '1'),
            ('holding', '1'),
            ('3', '1'),

            # Input Register (Info1: 4) -> Default to 4 (RO)
            ('Input Register', '4'),
            ('input', '4'),
            ('4', '4'),

            # Coil (Info1: 1) -> Default to 1 (RW)
            ('Coil', '1'),
            ('coils', '1'),
            ('1', '1'),

            # Discrete Input (Info1: 2) -> Default to 4 (RO)
            ('Discrete Input', '4'),
            ('2', '4'),

            # Default fallback (Unknown -> Holding -> 3) -> 1
            ('', '1'),
            (None, '1'),
            ('Unknown', '1'),
        ]

        for reg_type, expected_action in test_cases:
            rows = [{
                'Name': f'Test_{reg_type}',
                'Address': '100',
                'Type': 'U16',
                'RegisterType': reg_type,
                'Action': '' # Test defaulting
            }]
            processed = list(self.generator.process_rows(rows))
            with self.subTest(reg_type=reg_type):
                self.assertEqual(processed[0]['Action'], expected_action, f"Failed for {reg_type}")

    def test_explicit_action_preservation(self):
        # Even if it's an Input Register, if an explicit action is given, preserve it (if valid)
        rows = [{
            'Name': 'TestExplicit',
            'Address': '100',
            'Type': 'U16',
            'RegisterType': 'Input Register',
            'Action': '1' # Force RW on Input Register
        }]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

if __name__ == '__main__':
    unittest.main()
