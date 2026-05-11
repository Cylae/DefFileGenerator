import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestActionDefaults(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        # Disable logging for tests to keep output clean
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)

    def test_coil_default_action(self):
        rows = [{'Name': 'Test Coil', 'Address': '1', 'RegisterType': 'Coil', 'Type': 'U16'}]
        results = list(self.generator.process_rows(rows))
        self.assertEqual(results[0]['Action'], '1')
        self.assertEqual(results[0]['Info1'], '1')

    def test_discrete_input_default_action(self):
        rows = [{'Name': 'Test DI', 'Address': '100', 'RegisterType': 'Discrete Input', 'Type': 'U16'}]
        results = list(self.generator.process_rows(rows))
        self.assertEqual(results[0]['Action'], '4')
        self.assertEqual(results[0]['Info1'], '2')

    def test_holding_register_default_action(self):
        rows = [{'Name': 'Test Holding', 'Address': '1000', 'RegisterType': 'Holding Register', 'Type': 'U16'}]
        results = list(self.generator.process_rows(rows))
        self.assertEqual(results[0]['Action'], '1')
        self.assertEqual(results[0]['Info1'], '3')

    def test_input_register_default_action(self):
        rows = [{'Name': 'Test Input', 'Address': '5000', 'RegisterType': 'Input Register', 'Type': 'U16'}]
        results = list(self.generator.process_rows(rows))
        self.assertEqual(results[0]['Action'], '4')
        self.assertEqual(results[0]['Info1'], '4')

    def test_explicit_action_overrides_default(self):
        rows = [{'Name': 'Test Input', 'Address': '5000', 'RegisterType': 'Input Register', 'Type': 'U16', 'Action': '1'}]
        results = list(self.generator.process_rows(rows))
        self.assertEqual(results[0]['Action'], '1')

    def test_synonym_action_overrides_default(self):
        rows = [{'Name': 'Test Holding', 'Address': '1000', 'RegisterType': 'Holding Register', 'Type': 'U16', 'Action': 'RO'}]
        results = list(self.generator.process_rows(rows))
        self.assertEqual(results[0]['Action'], '4')

if __name__ == '__main__':
    unittest.main()
