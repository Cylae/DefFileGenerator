import unittest
import sys
import os
import csv
from unittest.mock import patch

class TestCliEntryPoints(unittest.TestCase):
    def setUp(self):
        self.csv_file = "test_cli_input.csv"
        with open(self.csv_file, 'w', encoding='utf-8') as f:
            f.write("Address,Name,Type\n100,Test1,U16\n")
        # Suppress logging during tests
        import logging
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        if os.path.exists(self.csv_file):
            os.remove(self.csv_file)
        if os.path.exists("test_mfg_test_model_definition.csv"):
            os.remove("test_mfg_test_model_definition.csv")
        if os.path.exists("test_out.csv"):
            os.remove("test_out.csv")
        import logging
        logging.disable(logging.NOTSET)

    def test_doc_to_webdyn_success(self):
        from doc_to_webdyn import main
        test_args = ["doc_to_webdyn.py", self.csv_file, "--manufacturer", "Test Mfg", "--model", "Test Model"]
        with patch.object(sys, 'argv', test_args):
            main()
            self.assertTrue(os.path.exists("test_mfg_test_model_definition.csv"))

    def test_def_file_gen_main_validate_success(self):
        from DefFileGenerator.main import main
        # Create a valid definition file
        def_file = "valid_def.csv"
        with open(def_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])

        test_args = ["main.py", "validate", def_file]
        try:
            with patch.object(sys, 'argv', test_args):
                main()
        finally:
            if os.path.exists(def_file):
                os.remove(def_file)

    def test_def_file_gen_main_validate_failure(self):
        from DefFileGenerator.main import main
        # Create an invalid definition file
        def_file = "invalid_def.csv"
        with open(def_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'MODEL', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']) # Invalid address

        test_args = ["main.py", "validate", def_file]
        try:
            with patch.object(sys, 'argv', test_args):
                with self.assertRaises(SystemExit) as cm:
                    main()
                self.assertEqual(cm.exception.code, 1)
        finally:
            if os.path.exists(def_file):
                os.remove(def_file)

    def test_def_file_gen_main_generate_template(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "generate", "-o", "test_out.csv", "--template"]
        with patch.object(sys, 'argv', test_args):
            main()
            self.assertTrue(os.path.exists("test_out.csv"))
            with open("test_out.csv", 'r') as f:
                content = f.read()
                self.assertIn("Name,Tag,RegisterType", content)

    def test_def_file_gen_main_generate_template_definition(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "generate", "-o", "test_out.csv", "--template", "--template-mode", "definition"]
        with patch.object(sys, 'argv', test_args):
            main()
            self.assertTrue(os.path.exists("test_out.csv"))
            with open("test_out.csv", 'r') as f:
                content = f.read()
                self.assertIn("#Index;Info1;Info2", content)

    def test_def_file_gen_main_extract(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "extract", self.csv_file, "-o", "test_out.csv"]
        with patch.object(sys, 'argv', test_args):
            main()
            self.assertTrue(os.path.exists("test_out.csv"))

    def test_def_file_gen_main_run(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "run", self.csv_file, "--manufacturer", "Test", "--model", "Test", "-o", "test_out.csv"]
        with patch.object(sys, 'argv', test_args):
            main()
            self.assertTrue(os.path.exists("test_out.csv"))

    def test_def_file_gen_main_generate(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "generate", self.csv_file, "--manufacturer", "Test", "--model", "Test", "-o", "test_out.csv"]
        with patch.object(sys, 'argv', test_args):
            main()
            self.assertTrue(os.path.exists("test_out.csv"))

    def test_main_pages_ignored(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "extract", self.csv_file, "--pages", "1, 2, 3"]
        with patch.object(sys, 'argv', test_args):
            main()

if __name__ == "__main__":
    unittest.main()
