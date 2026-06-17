import unittest
import os
import sys
import io
from DefFileGenerator.def_gen import run_generator, GeneratorConfig

class TestTemplate(unittest.TestCase):
    def setUp(self):
        self.template_file = "test_template.csv"

    def tearDown(self):
        if os.path.exists(self.template_file):
            os.remove(self.template_file)

    def test_generate_template_file(self):
        config = GeneratorConfig(output=self.template_file, template=True)
        run_generator(config)
        self.assertTrue(os.path.exists(self.template_file))
        with open(self.template_file, 'r') as f:
            content = f.read()
            self.assertIn("Name,Tag,RegisterType,Address,Type,Factor,Offset,Unit,Action,ScaleFactor", content)
            self.assertIn("Example Variable", content)

    def test_generate_template_stdout(self):
        config = GeneratorConfig(output=None, template=True)
        captured_output = io.StringIO()
        old_stdout = sys.stdout
        sys.stdout = captured_output
        try:
            run_generator(config)
        finally:
            sys.stdout = old_stdout

        content = captured_output.getvalue()
        self.assertIn("Name,Tag,RegisterType,Address,Type,Factor,Offset,Unit,Action,ScaleFactor", content)
        self.assertIn("Example Variable", content)

if __name__ == "__main__":
    unittest.main()
