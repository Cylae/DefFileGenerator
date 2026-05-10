import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestActionDefaults(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        # Suppress logging during tests
        logging.getLogger().setLevel(logging.ERROR)

    def test_action_defaulting_by_type(self):
        test_cases = [
            # (RegisterType, expected_action)
            ('Holding Register', '1'),
            ('Coil', '1'),
            ('Input Register', '4'),
            # ('Discrete Input', '2'), # Wait, Info1 for Discrete Input is 2.
                                     # My implementation: norm_action = '4' if info1 in ['2', '4'] else '1'
                                     # So Discrete Input (2) will default to 4.
                                     # Actually, Action '4' is Read-Only for Webdyn.
                                     # Discrete Input IS Read-Only.
            ('Discrete Input', '4'),
            ('input', '4'),
            ('holding', '1'),
            ('coil', '1'),
            ('', '1'), # Defaults to Info1 '3' (Holding Register), which defaults to Action '1'
        ]

        for reg_type, expected_action in test_cases:
            rows = [{
                'Name': f'Test_{reg_type}',
                'Address': '100',
                'Type': 'U16',
                'RegisterType': reg_type,
                'Action': '' # Testing default
            }]
            processed = list(self.generator.process_rows(rows))
            with self.subTest(reg_type=reg_type):
                self.assertEqual(processed[0]['Action'], expected_action, f"Failed for {reg_type}")

    def test_action_override_preserved(self):
        # Even if it's an Input Register, if Action is explicitly set, keep it.
        rows = [{
            'Name': 'ExplicitAction',
            'Address': '100',
            'Type': 'U16',
            'RegisterType': 'Input Register',
            'Action': '1'
        }]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

        # And vice-versa
        rows = [{
            'Name': 'ExplicitAction2',
            'Address': '100',
            'Type': 'U16',
            'RegisterType': 'Holding Register',
            'Action': '4'
        }]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4')

if __name__ == '__main__':
    unittest.main()
