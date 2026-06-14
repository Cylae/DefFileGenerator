import unittest
import os
import csv
from DefFileGenerator.def_gen import run_generator, GeneratorConfig

class TestTemplate(unittest.TestCase):
    def setUp(self):
        self.filename = "test_template.csv"

    def tearDown(self):
        if os.path.exists(self.filename):
            os.remove(self.filename)

    def test_generate_template(self):
        config = GeneratorConfig(output=self.filename, template=True)
        run_generator(config)

        self.assertTrue(os.path.exists(self.filename))
        with open(self.filename, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            headers = next(reader)
            self.assertIn("Name", headers)
            self.assertIn("Address", headers)

            row = next(reader)
            self.assertEqual(row[0], "Example Variable")

if __name__ == '__main__':
    unittest.main()
