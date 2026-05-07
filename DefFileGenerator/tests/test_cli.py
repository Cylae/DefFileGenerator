import unittest
import sys
import os
import io
import json
from unittest.mock import patch

class TestCliEntryPoints(unittest.TestCase):
    def setUp(self):
        self.csv_file = "test_cli_input.csv"
        with open(self.csv_file, 'w', encoding='utf-8') as f:
            f.write("Address,Name,Type\n100,Test1,U16\n")

    def tearDown(self):
        if os.path.exists(self.csv_file):
            os.remove(self.csv_file)
        if os.path.exists("test_mfg_test_model_definition.csv"):
            os.remove("test_mfg_test_model_definition.csv")
        if os.path.exists("test_out.csv"):
            os.remove("test_out.csv")

    def test_doc_to_webdyn_success(self):
        from doc_to_webdyn import main
        test_args = ["doc_to_webdyn.py", self.csv_file, "--manufacturer", "Test Mfg", "--model", "Test Model"]
        with patch.object(sys, 'argv', test_args):
            captured_output = io.StringIO()
            sys.stdout = captured_output
            try:
                main()
            finally:
                sys.stdout = sys.__stdout__
            self.assertTrue(os.path.exists("test_mfg_test_model_definition.csv"))

    def test_doc_to_webdyn_missing_file(self):
        from doc_to_webdyn import main
        test_args = ["doc_to_webdyn.py", "nonexistent.csv", "--manufacturer", "Test", "--model", "Test"]
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)

    def test_doc_to_webdyn_bad_mapping(self):
        from doc_to_webdyn import main
        test_args = ["doc_to_webdyn.py", self.csv_file, "--manufacturer", "Test", "--model", "Test", "--mapping", "nonexistent.json"]
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)

    def test_doc_to_webdyn_empty_data(self):
        from doc_to_webdyn import main
        empty_csv = "test_empty.csv"
        with open(empty_csv, 'w') as f: f.write("")
        test_args = ["doc_to_webdyn.py", empty_csv, "--manufacturer", "Test", "--model", "Test"]
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.remove(empty_csv)

    def test_doc_to_webdyn_main_exception(self):
        from doc_to_webdyn import main
        with patch('doc_to_webdyn._run_cli', side_effect=ValueError("Mock Exception")):
            with self.assertLogs(level='ERROR') as log:
                with self.assertRaises(SystemExit) as cm:
                    main()
                self.assertEqual(cm.exception.code, 1)
                self.assertTrue(any("An unexpected error occurred" in m for m in log.output))

    def test_doc_to_webdyn_main_interrupt(self):
        from doc_to_webdyn import main
        with patch('doc_to_webdyn._run_cli', side_effect=KeyboardInterrupt()):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 130)

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

    def test_doc_to_webdyn_pages(self):
        from doc_to_webdyn import main
        test_args = ["doc_to_webdyn.py", self.csv_file, "--manufacturer", "Test", "--model", "Test", "--pages", "1,2"]
        with patch.object(sys, 'argv', test_args):
            main() # Should warn about non-pdf and continue
            self.assertTrue(os.path.exists("test_test_definition.csv"))
            if os.path.exists("test_test_definition.csv"):
                os.remove("test_test_definition.csv")

    def test_def_file_gen_pages_invalid(self):
        from doc_to_webdyn import main
        test_args = ["doc_to_webdyn.py", "dummy.pdf", "--manufacturer", "Test", "--model", "Test", "--pages", "a,b"]
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)

    def test_def_file_gen_main_pages_invalid(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "extract", "dummy.pdf", "--pages", "a,b"]
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)

    def test_def_file_gen_main_interrupt(self):
        from DefFileGenerator.main import main
        with patch('DefFileGenerator.main._run_cli', side_effect=KeyboardInterrupt()):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 130)

    def test_def_file_gen_main_exception(self):
        from DefFileGenerator.main import main
        with patch('DefFileGenerator.main._run_cli', side_effect=ValueError("Mock")):
            with self.assertLogs(level='ERROR') as log:
                with self.assertRaises(SystemExit) as cm:
                    main()
                self.assertEqual(cm.exception.code, 1)
                self.assertTrue(any("unexpected error" in m for m in log.output))

    def test_doc_to_webdyn_empty_mapped(self):
        from doc_to_webdyn import main
        bad_csv = "test_cli_empty_mapped.csv"
        with open(bad_csv, "w") as f: f.write("Random1,Random2\n100,200\n")

        test_args = ["doc_to_webdyn.py", bad_csv, "--manufacturer", "Test", "--model", "Test"]
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.remove(bad_csv)

    def test_def_file_gen_main_bad_mapping(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "extract", self.csv_file, "--mapping", "nonexistent.json"]
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)

    def test_def_file_gen_main_missing_file(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "extract", "nonexistent.csv"]
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)

    def test_def_file_gen_main_unsupported_ext(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "extract", "test.txt"]
        with open("test.txt", "w") as f: f.write("Dummy")
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.remove("test.txt")

    def test_doc_to_webdyn_pages_invalid_pdf(self):
        from doc_to_webdyn import main
        test_args = ["doc_to_webdyn.py", "dummy.pdf", "--manufacturer", "Test", "--model", "Test", "--pages", "invalid"]
        with open("dummy.pdf", "w") as f: f.write("dummy")
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.remove("dummy.pdf")

    def test_doc_to_webdyn_unsupported_ext(self):
        from doc_to_webdyn import main
        test_args = ["doc_to_webdyn.py", "test.txt", "--manufacturer", "Test", "--model", "Test"]
        with open("test.txt", "w") as f: f.write("dummy")
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.remove("test.txt")

    def test_doc_to_webdyn_valid_mapping(self):
        from doc_to_webdyn import main
        mapping_file = "valid_mapping.json"
        with open(mapping_file, "w") as f: json.dump({"Address": "Address"}, f)
        test_args = ["doc_to_webdyn.py", self.csv_file, "--manufacturer", "Test", "--model", "Test", "--mapping", mapping_file]
        with patch.object(sys, 'argv', test_args):
            main()
        os.remove(mapping_file)
        if os.path.exists("test_test_definition.csv"):
            os.remove("test_test_definition.csv")

    def test_def_file_gen_main_no_command(self):
        from DefFileGenerator.main import main
        test_args = ["main.py"]
        with patch.object(sys, 'argv', test_args):
            main() # Just prints help and returns

    def test_def_file_gen_main_excel(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "extract", "dummy.xlsx", "--sheet", "Sheet1"]
        with open("dummy.xlsx", "w") as f: f.write("Dummy")
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.remove("dummy.xlsx")

    def test_def_file_gen_main_pdf(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "extract", "dummy.pdf", "--pages", "1"]
        with open("dummy.pdf", "w") as f: f.write("Dummy")
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.remove("dummy.pdf")

    def test_def_file_gen_main_xml(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "extract", "dummy.xml"]
        with open("dummy.xml", "w") as f: f.write("Dummy")
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.remove("dummy.xml")

    def test_doc_to_webdyn_excel(self):
        from doc_to_webdyn import main
        test_args = ["doc_to_webdyn.py", "dummy.xlsx", "--manufacturer", "Test", "--model", "Test"]
        with open("dummy.xlsx", "w") as f: f.write("Dummy")
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit):
                main()
        os.remove("dummy.xlsx")

    def test_doc_to_webdyn_pdf(self):
        from doc_to_webdyn import main
        test_args = ["doc_to_webdyn.py", "dummy.pdf", "--manufacturer", "Test", "--model", "Test", "--pages", "1"]
        with open("dummy.pdf", "w") as f: f.write("Dummy")
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit):
                main()
        os.remove("dummy.pdf")

    def test_doc_to_webdyn_xml(self):
        from doc_to_webdyn import main
        test_args = ["doc_to_webdyn.py", "dummy.xml", "--manufacturer", "Test", "--model", "Test"]
        with open("dummy.xml", "w") as f: f.write("Dummy")
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit):
                main()
        os.remove("dummy.xml")

    def test_doc_to_webdyn_pages_parsing(self):
        from doc_to_webdyn import main
        test_args = ["doc_to_webdyn.py", "dummy.pdf", "--manufacturer", "Test", "--model", "Test", "--pages", "1, 2, 3"]
        with open("dummy.pdf", "w") as f: f.write("Dummy")
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.remove("dummy.pdf")

    def test_main_pages_parsing(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "extract", "dummy.pdf", "--pages", "1, 2, 3"]
        with open("dummy.pdf", "w") as f: f.write("Dummy")
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.remove("dummy.pdf")

    def test_main_pages_ignored(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "extract", self.csv_file, "--pages", "1, 2, 3"]
        with patch.object(sys, 'argv', test_args):
            main()

    def test_doc_to_webdyn_pages_parsing_error(self):
        from doc_to_webdyn import main
        test_args = ["doc_to_webdyn.py", "dummy.pdf", "--manufacturer", "Test", "--model", "Test", "--pages", "a,b"]
        with open("dummy.pdf", "w") as f: f.write("Dummy")
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.remove("dummy.pdf")

    def test_main_pages_parsing_error(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "extract", "dummy.pdf", "--pages", "a,b"]
        with open("dummy.pdf", "w") as f: f.write("Dummy")
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.remove("dummy.pdf")

    def test_main_run_pages_parsing_error(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "run", "dummy.pdf", "--manufacturer", "A", "--model", "B", "--pages", "a,b"]
        with open("dummy.pdf", "w") as f: f.write("Dummy")
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.remove("dummy.pdf")

    def test_doc_to_webdyn_bad_mapping_open(self):
        from doc_to_webdyn import main
        test_args = ["doc_to_webdyn.py", self.csv_file, "--manufacturer", "Test", "--model", "Test", "--mapping", "dummy_dir"]
        os.mkdir("dummy_dir")
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.rmdir("dummy_dir")

    def test_main_bad_mapping_open(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "extract", self.csv_file, "--mapping", "dummy_dir"]
        os.mkdir("dummy_dir")
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.rmdir("dummy_dir")

    def test_def_file_gen_main_extract_stdout(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "extract", self.csv_file]
        with patch.object(sys, 'argv', test_args):
            captured_output = io.StringIO()
            sys.stdout = captured_output
            try:
                main()
            finally:
                sys.stdout = sys.__stdout__
            self.assertIn("Test1", captured_output.getvalue())

    def test_def_file_gen_main_run_no_data(self):
        from DefFileGenerator.main import main
        empty_csv = "test_empty.csv"
        with open(empty_csv, 'w') as f: f.write("")
        test_args = ["main.py", "run", empty_csv, "--manufacturer", "Test", "--model", "Test"]
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.remove(empty_csv)

    def test_def_file_gen_main_extract_bad_mapping_open(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "extract", self.csv_file, "--mapping", "nonexistent_mapping.json"]
        with patch.object(sys, 'argv', test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)

if __name__ == "__main__":
    unittest.main()
