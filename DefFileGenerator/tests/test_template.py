import unittest
import os
import csv
from DefFileGenerator.def_gen import run_generator, GeneratorConfig

class TestTemplate(unittest.TestCase):
    def setUp(self):
        self.template_file = "test_template.csv"

    def tearDown(self):
        if os.path.exists(self.template_file):
            os.remove(self.template_file)

    def test_generate_template(self):
        config = GeneratorConfig(output=self.template_file, template=True)
        run_generator(config)

        self.assertTrue(os.path.exists(self.template_file))

        with open(self.template_file, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            headers = next(reader)
            self.assertEqual(headers[0], 'Name')
            self.assertEqual(headers[3], 'Address')

            first_row = next(reader)
            self.assertEqual(first_row[0], 'Example Variable')
            self.assertEqual(first_row[3], '30001')

if __name__ == '__main__':
    unittest.main()
