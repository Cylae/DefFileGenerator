import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestActionDefaults(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        # Suppress logging warnings during tests
        logging.getLogger().setLevel(logging.ERROR)

    def test_holding_register_default_action(self):
        # RegisterType '3' or 'Holding Register' -> Action '1'
        rows = [{'Name': 'Test', 'Address': '1', 'RegisterType': 'Holding Register'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

    def test_input_register_default_action(self):
        # RegisterType '4' or 'Input Register' -> Action '4'
        rows = [{'Name': 'Test', 'Address': '1', 'RegisterType': 'Input Register'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4')

    def test_coil_default_action(self):
        # RegisterType '1' or 'Coil' -> Action '1'
        rows = [{'Name': 'Test', 'Address': '1', 'RegisterType': 'Coil'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

    def test_discrete_input_default_action(self):
        # RegisterType '2' or 'Discrete Input' -> Action '4'
        rows = [{'Name': 'Test', 'Address': '1', 'RegisterType': 'Discrete Input'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4')

    def test_explicit_action_preserved(self):
        # If Action is provided, it should be preserved (and normalized)
        rows = [{'Name': 'Test', 'Address': '1', 'RegisterType': 'Holding Register', 'Action': '4'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4')

if __name__ == '__main__':
    unittest.main()
