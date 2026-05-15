import unittest
from DefFileGenerator.def_gen import Generator

class TestActionDefaults(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_intelligent_action_defaults(self):
        """Verify that actions default based on register type if not provided."""
        test_cases = [
            # (RegisterType, Expected Action)
            ('Holding Register', '1'), # Info1='3' -> Read/Write
            ('Input Register', '4'),   # Info1='4' -> Read Only
            ('Coil', '1'),             # Info1='1' -> Read/Write
            ('Discrete Input', '4'),   # Info1='2' -> Read Only
        ]

        for reg_type, expected_action in test_cases:
            rows = [{
                'Name': f'Test_{reg_type}',
                'Address': '100',
                'Type': 'U16',
                'RegisterType': reg_type,
                'Action': ''  # Trigger defaulting
            }]
            processed = list(self.generator.process_rows(rows))
            with self.subTest(reg_type=reg_type):
                self.assertEqual(processed[0]['Action'], expected_action, f"Failed for {reg_type}")

    def test_explicit_action_overrides_default(self):
        rows = [{
            'Name': 'TestOverride',
            'Address': '100',
            'Type': 'U16',
            'RegisterType': 'Input Register',
            'Action': '1' # Explicitly set to R/W even if it's an Input Register
        }]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

if __name__ == '__main__':
    unittest.main()
