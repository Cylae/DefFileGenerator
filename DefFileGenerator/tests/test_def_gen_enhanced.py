import unittest
import logging
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator
import os
import csv
import tempfile

class TestGeneratorEnhanced(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_address_offset(self):
        rows = [{'Name': 'Test', 'Address': '100', 'Type': 'U16'}]
        processed = self.generator.process_rows(rows, address_offset=10)
        self.assertEqual(processed[0]['Info2'], '110')

        processed_neg = self.generator.process_rows(rows, address_offset=-110)
        self.assertEqual(processed_neg[0]['Info2'], '-10')

    def test_normalize_type_synonyms(self):
        self.assertEqual(self.generator.normalize_type('unsigned int 32'), 'U32')
        self.assertEqual(self.generator.normalize_type('signed integer 16_W'), 'I16_W')
        self.assertEqual(self.generator.normalize_type('float64'), 'F64')

    def test_negative_address_validation(self):
        self.assertTrue(self.generator.validate_address('-10', 'U16'))

    def test_run_generator_config(self):
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as tf:
            writer = csv.writer(tf)
            writer.writerow(['Name', 'RegisterType', 'Address', 'Type'])
            writer.writerow(['Var1', 'Holding Register', '100', 'U16'])
            temp_input = tf.name

        output_file = temp_input + "_out.csv"
        config = GeneratorConfig(
            input_file=temp_input,
            output=output_file,
            manufacturer='TestMFG',
            model='TestModel'
        )
        try:
            run_generator(config)
            self.assertTrue(os.path.exists(output_file))
            with open(output_file, 'r') as f:
                content = f.read()
                self.assertIn('TestMFG', content)
                self.assertIn('TestModel', content)
                self.assertIn('Var1', content)
        finally:
            if os.path.exists(temp_input): os.remove(temp_input)
            if os.path.exists(output_file): os.remove(output_file)

if __name__ == '__main__':
    unittest.main()
