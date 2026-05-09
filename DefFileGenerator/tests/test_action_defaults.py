import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestActionDefaults(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        # Suppress logging during tests
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)

    def test_action_defaults(self):
        test_rows = [
            {'Name': 'CoilVar', 'RegisterType': 'Coil', 'Address': '1', 'Type': 'U16'},
            {'Name': 'DI_Var', 'RegisterType': 'Discrete Input', 'Address': '2', 'Type': 'U16'},
            {'Name': 'HoldingVar', 'RegisterType': 'Holding Register', 'Address': '3', 'Type': 'U16'},
            {'Name': 'InputVar', 'RegisterType': 'Input Register', 'Address': '4', 'Type': 'U16'}
        ]

        processed = list(self.generator.process_rows(test_rows))

        # Mapping:
        # Coils (Info1=1) -> Action 1
        # Discrete Inputs (Info1=2) -> Action 4
        # Holding Registers (Info1=3) -> Action 1
        # Input Registers (Info1=4) -> Action 4

        self.assertEqual(processed[0]['Name'], 'CoilVar')
        self.assertEqual(processed[0]['Info1'], '1')
        self.assertEqual(processed[0]['Action'], '1')

        self.assertEqual(processed[1]['Name'], 'DI_Var')
        self.assertEqual(processed[1]['Info1'], '2')
        self.assertEqual(processed[1]['Action'], '4')

        self.assertEqual(processed[2]['Name'], 'HoldingVar')
        self.assertEqual(processed[2]['Info1'], '3')
        self.assertEqual(processed[2]['Action'], '1')

        self.assertEqual(processed[3]['Name'], 'InputVar')
        self.assertEqual(processed[3]['Info1'], '4')
        self.assertEqual(processed[3]['Action'], '4')

    def test_explicit_action_preserved(self):
        test_rows = [
            {'Name': 'ExplicitRO', 'RegisterType': 'Holding Register', 'Address': '10', 'Type': 'U16', 'Action': '4'},
            {'Name': 'ExplicitRW', 'RegisterType': 'Input Register', 'Address': '20', 'Type': 'U16', 'Action': '1'}
        ]

        processed = list(self.generator.process_rows(test_rows))

        self.assertEqual(processed[0]['Action'], '4')
        self.assertEqual(processed[1]['Action'], '1')

if __name__ == "__main__":
    unittest.main()
