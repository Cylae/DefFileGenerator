import unittest
from DefFileGenerator.def_gen import Generator

class TestActionDefaults(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_action_defaulting_by_type(self):
        test_cases = [
            # RegisterType, Expected Action (when input action is empty)
            ('Holding Register', '1'),
            ('Coil', '1'),
            ('Input Register', '4'),
            ('Discrete Input', '4'),
        ]

        for reg_type, expected_action in test_cases:
            rows = [{
                'Name': f'Test_{reg_type}',
                'Address': '100',
                'Type': 'U16',
                'RegisterType': reg_type,
                'Action': ''
            }]
            processed = list(self.generator.process_rows(rows))
            with self.subTest(reg_type=reg_type):
                self.assertEqual(processed[0]['Action'], expected_action, f"Failed for {reg_type}")

    def test_explicit_action_overrides_default(self):
        # Even for Input Register, an explicit action should be respected
        rows = [{
            'Name': 'ExplicitInput',
            'Address': '100',
            'Type': 'U16',
            'RegisterType': 'Input Register',
            'Action': '1'
        }]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

if __name__ == '__main__':
    unittest.main()
