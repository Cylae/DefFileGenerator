import unittest
import os
import tempfile
import csv
from DefFileGenerator.def_gen import run_generator, GeneratorConfig

class TestTemplate(unittest.TestCase):
    def test_generate_template(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            output_file = os.path.join(tmpdir, "template.csv")
            config = GeneratorConfig(template=True, output=output_file)
            run_generator(config)

            self.assertTrue(os.path.exists(output_file))
            with open(output_file, 'r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                header = next(reader)
                self.assertEqual(header[0], 'Name')
                self.assertEqual(header[1], 'Tag')

                rows = list(reader)
                self.assertGreater(len(rows), 0)
                self.assertEqual(rows[0][0], 'Example Variable')

if __name__ == '__main__':
    unittest.main()
