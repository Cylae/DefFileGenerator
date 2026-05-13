import unittest
from DefFileGenerator.def_gen import Generator

class TestActionDefaults(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_action_defaults(self):
        rows = [
            {'Name': 'Coil', 'Address': '1', 'RegisterType': 'Coil', 'Type': 'U16'},
            {'Name': 'DiscreteInput', 'Address': '10001', 'RegisterType': 'Discrete Input', 'Type': 'U16'},
            {'Name': 'HoldingRegister', 'Address': '40001', 'RegisterType': 'Holding Register', 'Type': 'U16'},
            {'Name': 'InputRegister', 'Address': '30001', 'RegisterType': 'Input Register', 'Type': 'U16'},
        ]

        processed = list(self.generator.process_rows(rows))

        self.assertEqual(processed[0]['Action'], '1') # Coil -> Read/Write
        self.assertEqual(processed[1]['Action'], '4') # Discrete Input -> Read Only
        self.assertEqual(processed[2]['Action'], '1') # Holding Register -> Read/Write
        self.assertEqual(processed[3]['Action'], '4') # Input Register -> Read Only

    def test_explicit_action_preserved(self):
        rows = [
            {'Name': 'ReadCoil', 'Address': '1', 'RegisterType': 'Coil', 'Type': 'U16', 'Action': '4'},
            {'Name': 'WriteInput', 'Address': '30001', 'RegisterType': 'Input Register', 'Type': 'U16', 'Action': '1'},
        ]

        processed = list(self.generator.process_rows(rows))

        self.assertEqual(processed[0]['Action'], '4')
        self.assertEqual(processed[1]['Action'], '1')

if __name__ == '__main__':
    unittest.main()
