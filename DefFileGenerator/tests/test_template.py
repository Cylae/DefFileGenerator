import unittest
import os
import csv
from DefFileGenerator.def_gen import generate_template

class TestTemplate(unittest.TestCase):
    def test_generate_template(self):
        filename = "test_template.csv"
        try:
            generate_template(filename)
            self.assertTrue(os.path.exists(filename))
            with open(filename, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                header = next(reader)
                self.assertEqual(header[0], 'Name')
                self.assertEqual(header[3], 'Address')

                row1 = next(reader)
                self.assertEqual(row1[0], 'Example Variable')
        finally:
            if os.path.exists(filename):
                os.remove(filename)

if __name__ == "__main__":
    unittest.main()
