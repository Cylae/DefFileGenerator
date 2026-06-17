import unittest
import os
import csv
from DefFileGenerator.def_gen import run_generator, GeneratorConfig

class TestTemplate(unittest.TestCase):
    def setUp(self):
        self.test_file = 'test_template.csv'

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def test_generate_template(self):
        config = GeneratorConfig(output=self.test_file, template=True)
        run_generator(config)

        self.assertTrue(os.path.exists(self.test_file))
        with open(self.test_file, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            header = next(reader)
            self.assertEqual(header[0], 'Name')
            self.assertEqual(header[3], 'Address')

            row1 = next(reader)
            self.assertEqual(row1[0], 'Example Variable')
            self.assertEqual(row1[3], '30001')

if __name__ == '__main__':
    unittest.main()
