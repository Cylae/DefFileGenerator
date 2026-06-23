import unittest
import os
import csv
from DefFileGenerator.def_gen import generate_template

class TestTemplate(unittest.TestCase):
    def setUp(self):
        self.template_file = "test_template.csv"

    def tearDown(self):
        if os.path.exists(self.template_file):
            os.remove(self.template_file)

    def test_generate_template(self):
        generate_template(self.template_file)
        self.assertTrue(os.path.exists(self.template_file))

        with open(self.template_file, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            headers = next(reader)
            self.assertEqual(headers, ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor'])

            row1 = next(reader)
            self.assertEqual(row1[0], 'Example Variable')
            self.assertEqual(row1[4], 'U16')

if __name__ == '__main__':
    unittest.main()
