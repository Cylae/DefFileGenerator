import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestActionDefaults(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        # Suppress logging during tests
        logging.getLogger().setLevel(logging.CRITICAL)

    def test_action_default_holding_register(self):
        rows = [{'Name': 'Test', 'Address': '100', 'Type': 'U16', 'RegisterType': 'Holding Register'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1') # Holding Register should default to Read-Write

    def test_action_default_input_register(self):
        rows = [{'Name': 'Test', 'Address': '100', 'Type': 'U16', 'RegisterType': 'Input Register'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4') # Input Register should default to Read-Only

    def test_action_default_coil(self):
        # BITS address format: Addr_Bit_Len
        rows = [{'Name': 'Test', 'Address': '100_0_1', 'Type': 'BITS', 'RegisterType': 'Coil'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1') # Coil should default to Read-Write

    def test_action_default_discrete_input(self):
        # BITS address format: Addr_Bit_Len
        rows = [{'Name': 'Test', 'Address': '100_0_1', 'Type': 'BITS', 'RegisterType': 'Discrete Input'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4') # Discrete Input should default to Read-Only

    def test_explicit_action_preserved(self):
        rows = [{'Name': 'Test', 'Address': '100', 'Type': 'U16', 'RegisterType': 'Input Register', 'Action': '1'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1') # Explicit action should be preserved

if __name__ == '__main__':
    unittest.main()
