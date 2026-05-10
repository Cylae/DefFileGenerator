import unittest
from DefFileGenerator.def_gen import Generator

class TestActionDefaults(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_action_defaulting(self):
        test_cases = [
            ('Coil', '1'), # Info1='1' -> RW default
            ('Discrete Input', '4'), # Info1='2' -> RO default
            ('Holding Register', '1'), # Info1='3' -> RW default
            ('Input Register', '4'), # Info1='4' -> RO default
            ('Holding', '1'), # Alias for Info1='3'
            ('Input', '4'), # Alias for Info1='4'
        ]

        for reg_type, expected_action in test_cases:
            rows = [{
                'Name': f'Test_{reg_type}',
                'Address': '100',
                'Type': 'U16',
                'RegisterType': reg_type,
                'Action': '' # Empty action to trigger defaulting
            }]
            processed = list(self.generator.process_rows(rows))
            with self.subTest(reg_type=reg_type):
                self.assertEqual(processed[0]['Action'], expected_action, f"Failed for {reg_type}")

    def test_explicit_action_overrides_default(self):
        # Even for Input Register, if '1' is provided, it should stay '1'
        rows = [{
            'Name': 'Explicit',
            'Address': '100',
            'Type': 'U16',
            'RegisterType': 'Input Register',
            'Action': '1'
        }]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

if __name__ == '__main__':
    unittest.main()
