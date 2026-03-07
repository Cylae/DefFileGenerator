import unittest
import logging
import os
import csv
import tempfile
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator

class TestGeneratorEnhanced(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_address_offset(self):
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
        # Process with offset 100
        processed = self.generator.process_rows(rows, address_offset=100)
        self.assertEqual(processed[0]['Info2'], '30100')

    def test_negative_address_warning(self):
        rows = [{
            'Name': 'Test Var',
            'Tag': 'test_tag',
            'RegisterType': 'Holding Register',
            'Address': '10',
            'Type': 'U16',
            'Factor': '1',
            'Offset': '0',
            'Unit': 'V',
            'Action': '4',
            'ScaleFactor': '0'
        }]
        # Process with offset -20 -> Resulting address -10
        with self.assertLogs(level='WARNING') as log:
            processed = self.generator.process_rows(rows, address_offset=-20)
            self.assertEqual(processed[0]['Info2'], '-10')
            self.assertTrue(any("Resulting address -10 is negative" in m for m in log.output))

    def test_normalize_type_synonyms(self):
        test_cases = [
            ('unsigned int 16', 'U16'),
            ('signed integer 32', 'I32'),
            ('float 64', 'F64'),
            ('uint16_w', 'U16_W'),
            ('int32_wb', 'I32_WB'),
            ('STR20', 'STR20'),
        ]
        for input_type, expected in test_cases:
            with self.subTest(input_type=input_type):
                self.assertEqual(self.generator.normalize_type(input_type), expected)

    def test_normalize_address_val_robust(self):
        # Memory says: RE_ADDR_VAL = re.compile(r'(?<![0-9A-Za-z])(0x[0-9A-Fa-f]+|[0-9A-Fa-f]+h|-?\d+|[0-9A-Fa-f]+)(?![0-9A-Za-z])')
        self.assertEqual(self.generator.normalize_address_val('Reg: 0x10'), '16')
        self.assertEqual(self.generator.normalize_address_val('Addr 7531h'), '30001')
        self.assertEqual(self.generator.normalize_address_val('-10'), '-10')
        self.assertEqual(self.generator.normalize_address_val('30,000'), '30000')

if __name__ == '__main__':
    unittest.main()
