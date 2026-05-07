import unittest
from DefFileGenerator.def_gen import Generator

class TestActionDefaults(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_action_defaults(self):
        rows = [
            {'Name': 'CoilVar', 'Address': '1', 'RegisterType': 'Coil', 'Type': 'U16'},
            {'Name': 'DiscIntVar', 'Address': '101', 'RegisterType': 'Discrete Input', 'Type': 'U16'},
            {'Name': 'HoldRegVar', 'Address': '40001', 'RegisterType': 'Holding Register', 'Type': 'U16'},
            {'Name': 'InputRegVar', 'Address': '30001', 'RegisterType': 'Input Register', 'Type': 'U16'},
        ]

        processed = list(self.generator.process_rows(rows))

        self.assertEqual(processed[0]['Action'], '1') # Coil -> RW (1)
        self.assertEqual(processed[1]['Action'], '4') # Discrete Input -> RO (4)
        self.assertEqual(processed[2]['Action'], '1') # Holding Register -> RW (1)
        self.assertEqual(processed[3]['Action'], '4') # Input Register -> RO (4)

    def test_explicit_action_preserved(self):
        rows = [
            {'Name': 'HoldRegRO', 'Address': '40001', 'RegisterType': 'Holding Register', 'Type': 'U16', 'Action': 'RO'},
            {'Name': 'InputRegRW', 'Address': '30001', 'RegisterType': 'Input Register', 'Type': 'U16', 'Action': 'RW'},
        ]

        processed = list(self.generator.process_rows(rows))

        self.assertEqual(processed[0]['Action'], '4') # Explicit RO
        self.assertEqual(processed[1]['Action'], '1') # Explicit RW

if __name__ == "__main__":
    unittest.main()
