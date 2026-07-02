#!/usr/bin/env python3
import unittest
import os
import tempfile
import csv
from DefFileGenerator.def_gen import run_generator, GeneratorConfig

class TestTemplate(unittest.TestCase):
    def test_template_generation(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            output_file = os.path.join(tmp_dir, 'template.csv')
            config = GeneratorConfig(output=output_file, template=True)
            run_generator(config)

            self.assertTrue(os.path.exists(output_file))
            with open(output_file, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                headers = next(reader)
                self.assertEqual(headers[0], 'Name')
                self.assertEqual(headers[3], 'Address')

                # Check at least one data row
                row1 = next(reader)
                self.assertEqual(row1[0], 'Example Variable')
                self.assertEqual(row1[3], '30001')

if __name__ == '__main__':
    unittest.main()
