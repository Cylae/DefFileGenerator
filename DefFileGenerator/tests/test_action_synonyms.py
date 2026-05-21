import unittest
from DefFileGenerator.def_gen import Generator

class TestActionSynonyms(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_action_normalization_extended(self):
        test_cases = [
            # Read-only synonyms -> 4
            ('R', '4'),
            ('READ', '4'),
            ('RO', '4'),
            ('READ-ONLY', '4'),
            ('READ ONLY', '4'),
            ('4', '4'),

            # Read/Write synonyms -> 1
            ('RW', '1'),
            ('W', '1'),
            ('WRITE', '1'),
            ('READ/WRITE', '1'),
            ('READ-WRITE', '1'),
            ('R/W', '1'),
            ('WO', '1'),
            ('WRITE-ONLY', '1'),
            ('WRITE ONLY', '1'),
            ('1', '1'),

            # Other allowed actions
            ('0', '0'),
            ('2', '2'),
            ('6', '6'),
            ('7', '7'),
            ('8', '8'),
            ('9', '9'),

            # Default/Fallback -> 1 (Defaults to 1 for Holding Registers)
            ('', '1'),
            ('UNKNOWN', '1'),
        ]

        for input_action, expected in test_cases:
            rows = [{
                'Name': 'TestVar',
                'Address': '100',
                'Type': 'U16',
                'Action': input_action
            }]
            processed = list(self.generator.process_rows(rows))
            with self.subTest(input_action=input_action):
                self.assertEqual(processed[0]['Action'], expected)

    def test_intelligent_action_defaulting(self):
        # Test that Input Register defaults to Action 4 (Read Only)
        rows = [{
            'Name': 'TestInput',
            'Address': '30001',
            'RegisterType': 'Input Register',
            'Type': 'U16',
            'Action': ''
        }]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4')

        # Test that Holding Register defaults to Action 1 (Read/Write)
        rows = [{
            'Name': 'TestHolding',
            'Address': '40001',
            'RegisterType': 'Holding Register',
            'Type': 'U16',
            'Action': ''
        }]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

if __name__ == '__main__':
    unittest.main()
