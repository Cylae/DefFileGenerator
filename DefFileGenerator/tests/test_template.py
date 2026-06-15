import unittest
import os
import tempfile
import csv
from DefFileGenerator.def_gen import run_generator, GeneratorConfig

class TestTemplateGeneration(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_generate_template(self):
        output_file = os.path.join(self.temp_dir.name, 'template.csv')
        config = GeneratorConfig(
            output=output_file,
            template=True
        )
        run_generator(config)

        self.assertTrue(os.path.exists(output_file))
        with open(output_file, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            header = next(reader)
            self.assertEqual(header[0], 'Name')
            self.assertEqual(header[3], 'Address')

            rows = list(reader)
            self.assertGreater(len(rows), 0)
            self.assertEqual(rows[0][0], 'Example Variable')

if __name__ == '__main__':
    unittest.main()
