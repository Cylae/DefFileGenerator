import unittest
import os
import csv
from DefFileGenerator.def_gen import generate_template

class TestTemplate(unittest.TestCase):
    def setUp(self):
        self.test_file = "test_template.csv"

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def test_generate_template_file(self):
        generate_template(self.test_file)
        self.assertTrue(os.path.exists(self.test_file))
        with open(self.test_file, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            headers = next(reader)
            self.assertEqual(headers[0], 'Name')
            self.assertEqual(headers[3], 'Address')

            row1 = next(reader)
            self.assertEqual(row1[0], 'Example Variable')

if __name__ == '__main__':
    unittest.main()
