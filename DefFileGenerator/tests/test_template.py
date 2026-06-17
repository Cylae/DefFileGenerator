import unittest
import os
from DefFileGenerator.def_gen import run_generator, GeneratorConfig

class TestTemplate(unittest.TestCase):
    def setUp(self):
        self.test_file = "test_template.csv"

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def test_template_generation(self):
        config = GeneratorConfig(output=self.test_file, template=True)
        run_generator(config)
        self.assertTrue(os.path.exists(self.test_file))
        with open(self.test_file, 'r') as f:
            content = f.read()
            self.assertIn("Name,Tag,RegisterType,Address,Type,Factor,Offset,Unit,Action,ScaleFactor", content)
            self.assertIn("Example Variable", content)

if __name__ == '__main__':
    unittest.main()
