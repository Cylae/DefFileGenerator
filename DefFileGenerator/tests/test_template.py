import unittest
import os
import csv
import tempfile
from DefFileGenerator.def_gen import generate_template, run_generator, GeneratorConfig

class TestTemplate(unittest.TestCase):
    def test_generate_template_to_file(self):
        with tempfile.NamedTemporaryFile(suffix='.csv', delete=False) as tmp:
            tmp_path = tmp.name

        try:
            generate_template(tmp_path)
            self.assertTrue(os.path.exists(tmp_path))
            with open(tmp_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                headers = next(reader)
                self.assertEqual(headers[0], 'Name')
                rows = list(reader)
                self.assertGreater(len(rows), 0)
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def test_run_generator_template_mode(self):
        with tempfile.NamedTemporaryFile(suffix='.csv', delete=False) as tmp:
            tmp_path = tmp.name

        try:
            config = GeneratorConfig(output=tmp_path, template=True)
            run_generator(config)
            self.assertTrue(os.path.exists(tmp_path))
            with open(tmp_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                headers = next(reader)
                self.assertEqual(headers[0], 'Name')
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

if __name__ == '__main__':
    unittest.main()
