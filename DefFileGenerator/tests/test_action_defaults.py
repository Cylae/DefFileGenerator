import unittest
from DefFileGenerator.def_gen import Generator

class TestActionDefaults(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_intelligent_action_defaulting(self):
        test_cases = [
            # (RegisterType, expected_action)
            ('Holding Register', '1'),
            ('holding', '1'),
            ('Input Register', '4'),
            ('input', '4'),
            ('Coil', '1'),
            ('Discrete Input', '4'),
            ('', '1'), # Defaults to Holding Register (3) -> 1
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
                self.assertEqual(processed[0]['Action'], expected_action, f"Failed for {reg_type}")

    def test_explicit_action_overrides_default(self):
        rows = [{
            'Name': 'TestVar',
            'Address': '100',
            'Type': 'U16',
            'RegisterType': 'Input Register',
            'Action': '1' # Explicit RW even if it's an Input Register
        }]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

if __name__ == '__main__':
    unittest.main()
