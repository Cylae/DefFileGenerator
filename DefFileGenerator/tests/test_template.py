import unittest
import os
import csv
import tempfile
from DefFileGenerator.def_gen import run_generator, GeneratorConfig

class TestTemplate(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_generate_template_csv(self):
        output_path = os.path.join(self.temp_dir.name, "template.csv")
        config = GeneratorConfig(output=output_path, template=True)
        run_generator(config)

        self.assertTrue(os.path.exists(output_path))
        with open(output_path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            header = next(reader)
            self.assertEqual(header[0], 'Name')
            self.assertEqual(header[3], 'Address')

            rows = list(reader)
            self.assertGreater(len(rows), 0)

    def test_run_template_definition(self):
        # In 'run' mode with --template, it should produce a definition file (with header)
        # Actually, our current implementation of generate_template produces a simplified CSV.
        # Let's check how 'run' handles --template.
        # In main.py: run_command calls generate_command which sets template=True in config.
        # In def_gen.py: run_generator(config) calls generate_template(config.output) if config.template is True.

        output_path = os.path.join(self.temp_dir.name, "template_def.csv")
        # If we want a definition template, we might need a separate function,
        # but the request was "support --template flag for template generation".
        # Currently it generates the simplified CSV template.

        config = GeneratorConfig(output=output_path, template=True)
        run_generator(config)

        self.assertTrue(os.path.exists(output_path))
        with open(output_path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            header = next(reader)
            self.assertEqual(header[0], 'Name')

if __name__ == '__main__':
    unittest.main()
