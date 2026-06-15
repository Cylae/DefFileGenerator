import unittest
import os
import csv
import tempfile
from DefFileGenerator.def_gen import generate_template

class TestTemplate(unittest.TestCase):
    def test_generate_template_to_file(self):
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.csv') as tmp:
            tmp_path = tmp.name

        try:
            generate_template(tmp_path)
            self.assertTrue(os.path.exists(tmp_path))
            with open(tmp_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                header = next(reader)
                self.assertEqual(header[0], 'Name')
                self.assertEqual(header[3], 'Address')

                rows = list(reader)
                self.assertGreaterEqual(len(rows), 2)
                self.assertEqual(rows[0][0], 'Example Variable')
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

if __name__ == '__main__':
    unittest.main()
