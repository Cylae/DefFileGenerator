import csv
import json
import os
import sys
import tempfile
import unittest
from unittest.mock import patch

from fastapi.testclient import TestClient

import DefFileGenerator.def_gen as def_gen
import DefFileGenerator.extractor as extractor
import generate_webdyn_def
from web.app import app


class TestCoverageWave13(unittest.TestCase):
    def setUp(self):
        self.client = TestClient(app)

    def test_generate_webdyn_def_main_demo(self):
        """Test generate_webdyn_def.main() when no args provided (demo mode)."""
        demo_in = "sample_register_map.csv"
        demo_out = "sample_output_definition.csv"

        # Cleanup if files exist beforehand
        for f in (demo_in, demo_out):
            if os.path.exists(f):
                os.remove(f)

        try:
            with patch.object(sys, "argv", ["generate_webdyn_def.py"]):
                with self.assertRaises(SystemExit) as cm:
                    generate_webdyn_def.main()
                self.assertEqual(cm.exception.code, 0)
        finally:
            for f in (demo_in, demo_out):
                if os.path.exists(f):
                    os.remove(f)

    def test_generate_webdyn_def_main_args(self):
        """Test generate_webdyn_def.main() with positional args."""
        with tempfile.TemporaryDirectory() as tmpdir:
            input_csv = os.path.join(tmpdir, "input.csv")
            output_csv = os.path.join(tmpdir, "output.csv")
            with open(input_csv, "w", encoding="utf-8") as f:
                f.write("Register,Name,Data Type,Unit,Scale,Access\n")
                f.write("40001,Power,uint16,W,1,R\n")

            test_argv = [
                "generate_webdyn_def.py",
                input_csv,
                output_csv,
                "TestMfg",
                "TestModel",
                "modbusRTU",
                "Inverter",
            ]
            with patch.object(sys, "argv", test_argv):
                with self.assertRaises(SystemExit) as cm:
                    generate_webdyn_def.main()
                self.assertEqual(cm.exception.code, 0)
            self.assertTrue(os.path.exists(output_csv))

    def test_web_app_read_root_no_index(self):
        """Test read_root() when index.html does not exist in static_dir."""
        with patch("web.app.static_dir", "/nonexistent_directory_path_12345"):
            response = self.client.get("/")
            self.assertEqual(response.status_code, 200)
            data = response.json()
            self.assertIn("message", data)

    def test_extractor_main_cli_csv(self):
        """Test Extractor main CLI entrypoint with CSV input and mapping."""
        with tempfile.TemporaryDirectory() as tmpdir:
            input_csv = os.path.join(tmpdir, "input.csv")
            output_csv = os.path.join(tmpdir, "output.csv")
            mapping_json = os.path.join(tmpdir, "mapping.json")

            with open(input_csv, "w", encoding="utf-8") as f:
                f.write("Reg,VarName,DType\n")
                f.write("100,Volt,uint16\n")

            with open(mapping_json, "w", encoding="utf-8") as f:
                json.dump({"Address": "Reg", "Name": "VarName", "Type": "DType"}, f)

            test_argv = [
                "extractor.py",
                input_csv,
                "-o",
                output_csv,
                "--mapping",
                mapping_json,
                "--address-offset",
                "10",
            ]
            with patch.object(sys, "argv", test_argv):
                extractor.main()

            self.assertTrue(os.path.exists(output_csv))
            with open(output_csv, encoding="utf-8") as f:
                reader = csv.DictReader(f)
                rows = list(reader)
                self.assertEqual(len(rows), 1)
                self.assertEqual(rows[0]["Address"], "110")

    def test_def_gen_main_template(self):
        """Test def_gen main CLI entrypoint with --template flag."""
        with tempfile.TemporaryDirectory() as tmpdir:
            output_csv = os.path.join(tmpdir, "template_def.csv")
            test_argv = [
                "def_gen.py",
                "-o",
                output_csv,
                "--template",
                "--template-mode",
                "definition",
            ]
            with patch.object(sys, "argv", test_argv):
                def_gen.main()

            self.assertTrue(os.path.exists(output_csv))
            with open(output_csv, encoding="utf-8") as f:
                content = f.read()
                self.assertIn("modbusRTU", content)


if __name__ == "__main__":
    unittest.main()
