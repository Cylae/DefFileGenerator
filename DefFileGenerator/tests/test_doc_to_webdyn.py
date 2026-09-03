import io
import json
import os
import sys
import unittest
from unittest.mock import patch


class TestDocToWebdyn(unittest.TestCase):
    def setUp(self):
        self.csv_file = "test_cli_input.csv"
        with open(self.csv_file, "w", encoding="utf-8") as f:
            f.write("Address,Name,Type\n100,Test1,U16\n")

    def tearDown(self):
        if os.path.exists(self.csv_file):
            os.remove(self.csv_file)
        if os.path.exists("test_mfg_test_model_definition.csv"):
            os.remove("test_mfg_test_model_definition.csv")
        if os.path.exists("test_test_definition.csv"):
            os.remove("test_test_definition.csv")

    def test_doc_to_webdyn_success(self):
        from doc_to_webdyn import main

        test_args = [
            "doc_to_webdyn.py",
            self.csv_file,
            "--manufacturer",
            "Test Mfg",
            "--model",
            "Test Model",
        ]
        with patch.object(sys, "argv", test_args):
            captured_output = io.StringIO()
            sys.stdout = captured_output
            try:
                main()
            finally:
                sys.stdout = sys.__stdout__
            self.assertTrue(os.path.exists("test_mfg_test_model_definition.csv"))

    def test_doc_to_webdyn_missing_file(self):
        from doc_to_webdyn import main

        test_args = [
            "doc_to_webdyn.py",
            "nonexistent.csv",
            "--manufacturer",
            "Test",
            "--model",
            "Test",
        ]
        with patch.object(sys, "argv", test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)

    def test_doc_to_webdyn_bad_mapping(self):
        from doc_to_webdyn import main

        test_args = [
            "doc_to_webdyn.py",
            self.csv_file,
            "--manufacturer",
            "Test",
            "--model",
            "Test",
            "--mapping",
            "nonexistent.json",
        ]
        with patch.object(sys, "argv", test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)

    def test_doc_to_webdyn_empty_data(self):
        from doc_to_webdyn import main

        empty_csv = "test_empty.csv"
        with open(empty_csv, "w") as f:
            f.write("")
        test_args = ["doc_to_webdyn.py", empty_csv, "--manufacturer", "Test", "--model", "Test"]
        with patch.object(sys, "argv", test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.remove(empty_csv)

    def test_doc_to_webdyn_main_exception(self):
        from doc_to_webdyn import main

        with patch("doc_to_webdyn._run_cli", side_effect=ValueError("Mock Exception")):
            with self.assertLogs(level="ERROR") as log:
                with self.assertRaises(SystemExit) as cm:
                    main()
                self.assertEqual(cm.exception.code, 1)
                self.assertTrue(any("An unexpected error occurred" in m for m in log.output))

    def test_doc_to_webdyn_main_interrupt(self):
        from doc_to_webdyn import main

        with patch("doc_to_webdyn._run_cli", side_effect=KeyboardInterrupt()):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 130)

    def test_doc_to_webdyn_pages(self):
        from doc_to_webdyn import main

        test_args = [
            "doc_to_webdyn.py",
            self.csv_file,
            "--manufacturer",
            "Test",
            "--model",
            "Test",
            "--pages",
            "1,2",
        ]
        with patch.object(sys, "argv", test_args):
            main()
            self.assertTrue(os.path.exists("test_test_definition.csv"))
            if os.path.exists("test_test_definition.csv"):
                os.remove("test_test_definition.csv")

    def test_doc_to_webdyn_empty_mapped(self):
        from doc_to_webdyn import main

        bad_csv = "test_cli_empty_mapped.csv"
        with open(bad_csv, "w") as f:
            f.write("Random1,Random2\n100,200\n")
        test_args = ["doc_to_webdyn.py", bad_csv, "--manufacturer", "Test", "--model", "Test"]
        with patch.object(sys, "argv", test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.remove(bad_csv)

    def test_doc_to_webdyn_pages_invalid_pdf(self):
        from doc_to_webdyn import main

        test_args = [
            "doc_to_webdyn.py",
            "dummy.pdf",
            "--manufacturer",
            "Test",
            "--model",
            "Test",
            "--pages",
            "invalid",
        ]
        with open("dummy.pdf", "w") as f:
            f.write("dummy")
        with patch.object(sys, "argv", test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.remove("dummy.pdf")

    def test_doc_to_webdyn_unsupported_ext(self):
        from doc_to_webdyn import main

        test_args = ["doc_to_webdyn.py", "test.txt", "--manufacturer", "Test", "--model", "Test"]
        with open("test.txt", "w") as f:
            f.write("dummy")
        with patch.object(sys, "argv", test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.remove("test.txt")

    def test_doc_to_webdyn_valid_mapping(self):
        from doc_to_webdyn import main

        mapping_file = "valid_mapping.json"
        with open(mapping_file, "w") as f:
            json.dump({"Address": "Address"}, f)
        test_args = [
            "doc_to_webdyn.py",
            self.csv_file,
            "--manufacturer",
            "Test",
            "--model",
            "Test",
            "--mapping",
            mapping_file,
        ]
        with patch.object(sys, "argv", test_args):
            main()
        os.remove(mapping_file)
        if os.path.exists("test_test_definition.csv"):
            os.remove("test_test_definition.csv")

    def test_doc_to_webdyn_excel(self):
        from doc_to_webdyn import main

        test_args = ["doc_to_webdyn.py", "dummy.xlsx", "--manufacturer", "Test", "--model", "Test"]
        with open("dummy.xlsx", "w") as f:
            f.write("Dummy")
        with patch.object(sys, "argv", test_args):
            with self.assertRaises(SystemExit):
                main()
        os.remove("dummy.xlsx")

    def test_doc_to_webdyn_pdf(self):
        from doc_to_webdyn import main

        test_args = [
            "doc_to_webdyn.py",
            "dummy.pdf",
            "--manufacturer",
            "Test",
            "--model",
            "Test",
            "--pages",
            "1",
        ]
        with open("dummy.pdf", "w") as f:
            f.write("Dummy")
        with patch.object(sys, "argv", test_args):
            with self.assertRaises(SystemExit):
                main()
        os.remove("dummy.pdf")

    def test_doc_to_webdyn_xml(self):
        from doc_to_webdyn import main

        test_args = ["doc_to_webdyn.py", "dummy.xml", "--manufacturer", "Test", "--model", "Test"]
        with open("dummy.xml", "w") as f:
            f.write("Dummy")
        with patch.object(sys, "argv", test_args):
            with self.assertRaises(SystemExit):
                main()
        os.remove("dummy.xml")

    def test_doc_to_webdyn_pages_parsing(self):
        from doc_to_webdyn import main

        test_args = [
            "doc_to_webdyn.py",
            "dummy.pdf",
            "--manufacturer",
            "Test",
            "--model",
            "Test",
            "--pages",
            "1, 2, 3",
        ]
        with open("dummy.pdf", "w") as f:
            f.write("Dummy")
        with patch.object(sys, "argv", test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.remove("dummy.pdf")

    def test_doc_to_webdyn_pages_parsing_error(self):
        from doc_to_webdyn import main

        test_args = [
            "doc_to_webdyn.py",
            "dummy.pdf",
            "--manufacturer",
            "Test",
            "--model",
            "Test",
            "--pages",
            "a,b",
        ]
        with open("dummy.pdf", "w") as f:
            f.write("Dummy")
        with patch.object(sys, "argv", test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.remove("dummy.pdf")

    def test_doc_to_webdyn_bad_mapping_open(self):
        from doc_to_webdyn import main

        test_args = [
            "doc_to_webdyn.py",
            self.csv_file,
            "--manufacturer",
            "Test",
            "--model",
            "Test",
            "--mapping",
            "dummy_dir",
        ]
        os.mkdir("dummy_dir")
        with patch.object(sys, "argv", test_args):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
        os.rmdir("dummy_dir")

    def test_module_execution(self):
        import runpy

        test_args = [
            "doc_to_webdyn.py",
            self.csv_file,
            "--manufacturer",
            "Test Mfg",
            "--model",
            "Test Model",
        ]
        with patch.object(sys, "argv", test_args):
            captured_output = io.StringIO()
            sys.stdout = captured_output
            try:
                runpy.run_module("doc_to_webdyn", run_name="__main__")
            finally:
                sys.stdout = sys.__stdout__
            self.assertTrue(os.path.exists("test_mfg_test_model_definition.csv"))


if __name__ == "__main__":
    unittest.main()
