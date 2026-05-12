import unittest
from DefFileGenerator.def_gen import Generator

class TestActionDefaults(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_action_defaults(self):
        # Case 1: Input Register (Info1='4') -> Default to '4' (Read Only)
        rows = [{'Name': 'InputReg', 'Address': '1', 'Type': 'U16', 'RegisterType': 'input register'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4')

        # Case 2: Discrete Input (Info1='2') -> Default to '4' (Read Only)
        rows = [{'Name': 'DiscreteIn', 'Address': '1', 'Type': 'U16', 'RegisterType': 'discrete input'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4')

        # Case 3: Holding Register (Info1='3') -> Default to '1' (Read/Write)
        rows = [{'Name': 'HoldingReg', 'Address': '1', 'Type': 'U16', 'RegisterType': 'holding register'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

        # Case 4: Coil (Info1='1') -> Default to '1' (Read/Write)
        rows = [{'Name': 'Coil', 'Address': '1', 'Type': 'U16', 'RegisterType': 'coil'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

        # Case 5: Explicit Action provided -> Should be preserved (or normalized)
        rows = [{'Name': 'Explicit', 'Address': '1', 'Type': 'U16', 'RegisterType': 'input register', 'Action': '1'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

        rows = [{'Name': 'ExplicitRO', 'Address': '1', 'Type': 'U16', 'RegisterType': 'holding register', 'Action': 'RO'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4')

if __name__ == '__main__':
    unittest.main()
