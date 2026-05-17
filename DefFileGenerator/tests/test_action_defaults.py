import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestActionDefaults(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        # Suppress logging for tests
        logging.getLogger().setLevel(logging.ERROR)

    def test_action_defaults(self):
        # Data with missing Action
        rows = [
            {'name': 'Coil 1', 'address': '1', 'registertype': 'Coil', 'type': 'U16'},
            {'name': 'DI 1', 'address': '2', 'registertype': 'Discrete Input', 'type': 'U16'},
            {'name': 'HR 1', 'address': '3', 'registertype': 'Holding Register', 'type': 'U16'},
            {'name': 'IR 1', 'address': '4', 'registertype': 'Input Register', 'type': 'U16'},
        ]

        processed = list(self.generator.process_rows(rows))

        self.assertEqual(len(processed), 4)

        # Coil -> Info1='1', default Action='1' (RW)
        self.assertEqual(processed[0]['Info1'], '1')
        self.assertEqual(processed[0]['Action'], '1')

        # Discrete Input -> Info1='2', default Action='4' (RO)
        self.assertEqual(processed[1]['Info1'], '2')
        self.assertEqual(processed[1]['Action'], '4')

        # Holding Register -> Info1='3', default Action='1' (RW)
        self.assertEqual(processed[2]['Info1'], '3')
        self.assertEqual(processed[2]['Action'], '1')

        # Input Register -> Info1='4', default Action='4' (RO)
        self.assertEqual(processed[3]['Info1'], '4')
        self.assertEqual(processed[3]['Action'], '4')

    def test_explicit_action_preserved(self):
        rows = [
            {'name': 'IR 1', 'address': '4', 'registertype': 'Input Register', 'type': 'U16', 'action': '1'},
            {'name': 'HR 1', 'address': '3', 'registertype': 'Holding Register', 'type': 'U16', 'action': '4'},
        ]
        processed = list(self.generator.process_rows(rows))

        self.assertEqual(processed[0]['Action'], '1')
        self.assertEqual(processed[1]['Action'], '4')

if __name__ == '__main__':
    unittest.main()
