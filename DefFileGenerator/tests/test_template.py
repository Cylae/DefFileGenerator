import unittest
import os
import csv
from DefFileGenerator.def_gen import generate_template

class TestTemplate(unittest.TestCase):
    def setUp(self):
        self.template_csv = 'test_template.csv'

    def tearDown(self):
        if os.path.exists(self.template_csv):
            os.remove(self.template_csv)

    def test_generate_template(self):
        generate_template(self.template_csv)
        self.assertTrue(os.path.exists(self.template_csv))

        with open(self.template_csv, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            rows = list(reader)

            self.assertEqual(len(rows), 2)
            self.assertEqual(rows[0]['Name'], 'Example Variable')
            self.assertEqual(rows[1]['Type'], 'STR20')

            # Check headers
            expected_headers = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
            self.assertEqual(list(rows[0].keys()), expected_headers)

if __name__ == '__main__':
    unittest.main()
