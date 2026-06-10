import unittest
import sys
import os
import io
from unittest.mock import patch

class TestTemplateGeneration(unittest.TestCase):
    def tearDown(self):
        if os.path.exists("template_main.csv"):
            os.remove("template_main.csv")
        if os.path.exists("template_doc.csv"):
            os.remove("template_doc.csv")

    def test_main_template_generation(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "generate", "--template", "-o", "template_main.csv"]
        with patch.object(sys, 'argv', test_args):
            main()
            self.assertTrue(os.path.exists("template_main.csv"))
            with open("template_main.csv", "r") as f:
                content = f.read()
                self.assertIn("Name,Tag,RegisterType,Address,Type", content)

    def test_doc_to_webdyn_template_generation(self):
        from doc_to_webdyn import main
        test_args = ["doc_to_webdyn.py", "--template", "-o", "template_doc.csv"]
        with patch.object(sys, 'argv', test_args):
            main()
            self.assertTrue(os.path.exists("template_doc.csv"))
            with open("template_doc.csv", "r") as f:
                content = f.read()
                # Generator.generate_template produces simplified CSV headers
                self.assertIn("Name,Tag,RegisterType,Address,Type", content)

if __name__ == '__main__':
    unittest.main()
