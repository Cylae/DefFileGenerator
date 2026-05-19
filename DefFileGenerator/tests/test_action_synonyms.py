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

            # Default/Fallback (specified Holding) -> 1
            ('', '1'),
            ('UNKNOWN', '1'),
        ]

        for input_action, expected in test_cases:
            rows = [{
                'Name': 'TestVar',
                'Address': '100',
                'Type': 'U16',
                'Action': input_action,
                'RegisterType': 'Holding Register'
            }]
            processed = list(self.generator.process_rows(rows))
            with self.subTest(input_action=input_action):
                self.assertEqual(processed[0]['Action'], expected)

    def test_intelligent_action_defaulting(self):
        test_cases = [
            ('Holding Register', '1'),
            ('Coil', '1'),
            ('Input Register', '4'),
            ('Discrete Input', '4'),
            ('', '1'), # Default to Holding -> 1
        ]

        for reg_type, expected in test_cases:
            rows = [{
                'Name': 'TestVar',
                'Address': '100',
                'Type': 'U16',
                'Action': '',
                'RegisterType': reg_type
            }]
            processed = list(self.generator.process_rows(rows))
            with self.subTest(reg_type=reg_type):
                self.assertEqual(processed[0]['Action'], expected)

if __name__ == '__main__':
    unittest.main()
