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

    def test_generate_template(self):
        path = os.path.join(self.test_dir.name, "template.csv")
        generate_template(path)

        self.assertTrue(os.path.exists(path))
        with open(path, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
            self.assertGreater(len(rows), 0)
            self.assertIn('Name', reader.fieldnames)
            self.assertIn('Address', reader.fieldnames)
            self.assertEqual(rows[0]['Name'], 'Example Variable')

if __name__ == '__main__':
    unittest.main()
