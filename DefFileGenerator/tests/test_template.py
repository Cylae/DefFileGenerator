import unittest
import os
import tempfile
import csv
from DefFileGenerator.def_gen import run_generator, GeneratorConfig

class TestTemplate(unittest.TestCase):
    def test_template_generation(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            out_file = os.path.join(tmpdir, "template.csv")
            config = GeneratorConfig(output=out_file, template=True)
            run_generator(config)

            self.assertTrue(os.path.exists(out_file))
            with open(out_file, 'r', newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                headers = reader.fieldnames
                self.assertIn('Name', headers)
                self.assertIn('Address', headers)
                rows = list(reader)
                self.assertGreater(len(rows), 0)
                self.assertEqual(rows[0]['Name'], 'Example Variable')

if __name__ == '__main__':
    unittest.main()
