import unittest
import os
import sys
import io
from DefFileGenerator.main import main as main_cli
from doc_to_webdyn import main as doc_to_webdyn_main
from unittest.mock import patch

class TestTemplateGeneration(unittest.TestCase):
    def tearDown(self):
        for f in ["template_main.csv", "template_doc.csv", "Manufacturer_Model_definition.csv"]:
            if os.path.exists(f):
                os.remove(f)

    def test_main_cli_generate_template(self):
        test_args = ["main.py", "generate", "--template", "-o", "template_main.csv"]
        with patch.object(sys, 'argv', test_args):
            main_cli()
        self.assertTrue(os.path.exists("template_main.csv"))
        with open("template_main.csv", 'r') as f:
            content = f.read()
            self.assertIn("Name,Tag,RegisterType,Address,Type", content)

    def test_main_cli_run_template(self):
        test_args = ["main.py", "run", "--template", "-o", "template_main.csv"]
        with patch.object(sys, 'argv', test_args):
            main_cli()
        self.assertTrue(os.path.exists("template_main.csv"))

    def test_doc_to_webdyn_template(self):
        test_args = ["doc_to_webdyn.py", "--template", "-o", "template_doc.csv"]
        with patch.object(sys, 'argv', test_args):
            doc_to_webdyn_main()
        self.assertTrue(os.path.exists("template_doc.csv"))

    def test_main_cli_validate_success(self):
        def_file = "test_valid_def.csv"
        with open(def_file, "w") as f:
            f.write("modbusRTU;Inverter;Mfg;Model;;;;;;;\n")
            f.write("1;3;100;U16;;Var1;tag1;1.0;0.0;V;1\n")

        test_args = ["main.py", "validate", def_file]
        try:
            with patch.object(sys, 'argv', test_args):
                main_cli()
        finally:
            if os.path.exists(def_file):
                os.remove(def_file)

    def test_main_cli_validate_fail(self):
        def_file = "test_invalid_def.csv"
        with open(def_file, "w") as f:
            f.write("modbusRTU;Inverter;Mfg;Model;;;;;;;\n")
            f.write("1;3;100;U16;;Var1;dup;1.0;0.0;V;1\n")
            f.write("2;3;101;U16;;Var2;dup;1.0;0.0;V;1\n")

        test_args = ["main.py", "validate", def_file]
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main_cli()
            self.assertEqual(cm.exception.code, 1)
        if os.path.exists(def_file):
            os.remove(def_file)

if __name__ == "__main__":
    unittest.main()
