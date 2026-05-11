import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestActionDefaults(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        # Suppress logging during tests
        logging.getLogger().setLevel(logging.ERROR)

    def test_action_defaults_for_input_registers(self):
        # Info1 '4' -> Input Register, should default to '4' (Read Only)
        rows = [{'Name': 'InputVar', 'Address': '100', 'RegisterType': 'Input Register', 'Type': 'U16'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4')

    def test_action_defaults_for_discrete_inputs(self):
        # Info1 '2' -> Discrete Input, should default to '4' (Read Only)
        rows = [{'Name': 'DiscVar', 'Address': '101', 'RegisterType': 'Discrete Input', 'Type': 'U16'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4')

    def test_action_defaults_for_holding_registers(self):
        # Info1 '3' -> Holding Register, should default to '1' (Read/Write)
        rows = [{'Name': 'HoldVar', 'Address': '102', 'RegisterType': 'Holding Register', 'Type': 'U16'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

    def test_action_defaults_for_coils(self):
        # Info1 '1' -> Coil, should default to '1' (Read/Write)
        rows = [{'Name': 'CoilVar', 'Address': '103', 'RegisterType': 'Coil', 'Type': 'U16'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

    def test_action_provided_overrides_default(self):
        # Even for Input Register, if '1' is provided, it should be kept
        rows = [{'Name': 'InputVar', 'Address': '104', 'RegisterType': 'Input Register', 'Type': 'U16', 'Action': '1'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

if __name__ == '__main__':
    unittest.main()
