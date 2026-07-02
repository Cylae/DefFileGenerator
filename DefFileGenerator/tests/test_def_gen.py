from unittest.mock import patch
from DefFileGenerator.def_gen import run_generator, GeneratorConfig
import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestGenerator(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        # Suppress logging during tests unless checking for logs
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)

    def test_intelligent_defaulting(self):
        rows = [
            {'Name': 'Holding', 'RegisterType': 'Holding', 'Address': '100', 'Type': 'U16'},
            {'Name': 'Input', 'RegisterType': 'Input', 'Address': '101', 'Type': 'U16'},
            {'Name': 'Coil', 'RegisterType': 'Coil', 'Address': '1', 'Type': 'U16'},
            {'Name': 'Discrete', 'RegisterType': 'Discrete', 'Address': '2', 'Type': 'U16'}
        ]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1') # Holding -> Read/Write
        self.assertEqual(processed[1]['Action'], '4') # Input -> Read Only
        self.assertEqual(processed[2]['Action'], '1') # Coil -> Read/Write
        self.assertEqual(processed[3]['Action'], '4') # Discrete -> Read Only

    def test_validate_address_range(self):
        self.assertTrue(self.generator.validate_address('0', 'U16'))
        self.assertTrue(self.generator.validate_address('65535', 'U16'))

        logging.disable(logging.NOTSET)
        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(self.generator.validate_address('65536', 'U16'))
            self.assertTrue(any("out of standard Modbus range" in m for m in log.output))
        logging.disable(logging.CRITICAL)

    def test_normalize_address_val(self):
        self.assertEqual(self.generator.normalize_address_val('0x10'), '16')
        self.assertEqual(self.generator.normalize_address_val('10h'), '16')
        self.assertEqual(self.generator.normalize_address_val('10'), '10')
        self.assertEqual(self.generator.normalize_address_val('A0'), '160')
        self.assertEqual(self.generator.normalize_address_val('1,234'), '1234')

    def test_process_rows_basic(self):
        rows = [{
            'Name': 'Test Var',
            'Tag': 'test_tag',
            'RegisterType': 'Holding Register',
            'Address': '30000',
            'Type': 'U16',
            'Factor': '1',
            'Offset': '0',
            'Unit': 'V',
            'Action': '4',
            'ScaleFactor': '0'
        }]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(len(processed), 1)
        self.assertEqual(processed[0]['Info1'], '3')
        self.assertEqual(processed[0]['Info3'], 'U16')
        self.assertEqual(processed[0]['CoefA'], '1.000000')

    def test_automatic_tag_generation(self):
        rows = [
            {'Name': 'Test Variable', 'Tag': '', 'RegisterType': '3', 'Address': '100', 'Type': 'U16'},
            {'Name': 'Test Variable', 'Tag': '', 'RegisterType': '3', 'Address': '101', 'Type': 'U16'}
        ]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Tag'], 'test_variable')
        self.assertEqual(processed[1]['Tag'], 'test_variable_1')

if __name__ == '__main__':
    unittest.main()
