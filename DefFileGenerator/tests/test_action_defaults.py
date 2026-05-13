import unittest
from DefFileGenerator.def_gen import Generator

class TestActionDefaults(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_action_defaulting(self):
        test_cases = [
            # (RegisterType, expected_action)
            ('Holding Register', '1'),
            ('Coil', '1'),
            ('Input Register', '4'),
            ('Discrete Input', '4'),
            ('', '1'), # Default reg type is Holding Register
            ('Unknown', '1'),
        ]

        for reg_type, expected in test_cases:
            rows = [{
                'Name': 'TestVar',
                'Address': '100',
                'Type': 'U16',
                'RegisterType': reg_type,
                'Action': '' # Unspecified action
            }]
            processed = list(self.generator.process_rows(rows))
            with self.subTest(reg_type=reg_type):
                self.assertEqual(processed[0]['Action'], expected)

    def test_explicit_action_overrides_default(self):
        rows = [{
            'Name': 'TestVar',
            'Address': '100',
            'Type': 'U16',
            'RegisterType': 'Input Register',
            'Action': '1' # Explicitly RW
        }]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

if __name__ == '__main__':
    unittest.main()
