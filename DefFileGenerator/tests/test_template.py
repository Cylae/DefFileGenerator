import unittest
import os
import csv
import tempfile
from DefFileGenerator.def_gen import generate_template

class TestTemplate(unittest.TestCase):
    def setUp(self):
        self.test_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.test_dir.cleanup()

    def test_generate_template_to_file(self):
        path = os.path.join(self.test_dir.name, "template.csv")
        generate_template(path)

        self.assertTrue(os.path.exists(path))
        with open(path, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            headers = next(reader)
            self.assertEqual(headers[0], 'Name')
            self.assertEqual(headers[1], 'Tag')

            row1 = next(reader)
            self.assertEqual(row1[0], 'Example Variable')

if __name__ == '__main__':
    unittest.main()
