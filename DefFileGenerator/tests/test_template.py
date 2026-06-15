import unittest
import os
import csv
from DefFileGenerator.def_gen import generate_template

class TestTemplate(unittest.TestCase):
    def setUp(self):
        self.test_file = 'test_template.csv'

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def test_generate_template(self):
        generate_template(self.test_file)
        self.assertTrue(os.path.exists(self.test_file))

        with open(self.test_file, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            header = next(reader)
            self.assertEqual(header[0], 'Name')
            self.assertEqual(header[1], 'Tag')

            rows = list(reader)
            self.assertGreater(len(rows), 0)
            self.assertEqual(rows[0][0], 'Example Variable')

if __name__ == '__main__':
    unittest.main()
