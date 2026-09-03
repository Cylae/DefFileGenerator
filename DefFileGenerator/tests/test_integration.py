import csv
import json
import os
import sys
import unittest
from unittest.mock import patch

from DefFileGenerator.main import main


class TestDefFileGeneratorIntegration(unittest.TestCase):
    def setUp(self):
        self.dummy_input = "integration_input.csv"
        self.mapping_file = "integration_mapping.json"
        self.output_extract = "integration_extract_out.csv"
        self.output_final = "integration_final_out.csv"

        # Create a dummy input that needs mapping
        with open(self.dummy_input, "w", encoding="utf-8") as f:
            f.write("RegAddr,VarName,RegType\n40001,Power,U32\n40003,Energy,U64\n")

        # Create mapping file
        mapping = {"Address": "RegAddr", "Name": "VarName", "Type": "RegType"}
        with open(self.mapping_file, "w", encoding="utf-8") as f:
            json.dump(mapping, f)

    def tearDown(self):
        for f in [self.dummy_input, self.mapping_file, self.output_extract, self.output_final]:
            if os.path.exists(f):
                os.remove(f)

    def test_end_to_end_extract_and_generate(self):
        # 1. Extract
        test_args_extract = [
            "main.py",
            "extract",
            self.dummy_input,
            "--mapping",
            self.mapping_file,
            "-o",
            self.output_extract,
        ]
        with patch.object(sys, "argv", test_args_extract):
            main()

        self.assertTrue(os.path.exists(self.output_extract))

        # Verify extracted content
        with open(self.output_extract, encoding="utf-8") as f:
            reader = csv.DictReader(f)
            rows = list(reader)
            self.assertEqual(len(rows), 2)
            self.assertEqual(rows[0]["Name"], "Power")
            self.assertEqual(rows[0]["Address"], "40001")
            self.assertEqual(rows[0]["Type"], "U32")

        # 2. Generate
        test_args_generate = [
            "main.py",
            "generate",
            self.output_extract,
            "--manufacturer",
            "TestMfg",
            "--model",
            "TestModel",
            "-o",
            self.output_final,
        ]
        with patch.object(sys, "argv", test_args_generate):
            main()

        self.assertTrue(os.path.exists(self.output_final))

        # Verify generated Webdyn definition content
        with open(self.output_final, encoding="utf-8") as f:
            reader = csv.reader(f, delimiter=";")
            rows = list(reader)
            self.assertTrue(len(rows) > 1)
            # Header row
            self.assertEqual(rows[0][2], "TestMfg")
            self.assertEqual(rows[0][3], "TestModel")
            # First register row
            self.assertEqual(rows[1][5], "Power")

    def test_end_to_end_run(self):
        # 1. Run (Extract + Generate)
        test_args_run = [
            "main.py",
            "run",
            self.dummy_input,
            "--mapping",
            self.mapping_file,
            "--manufacturer",
            "RunMfg",
            "--model",
            "RunModel",
            "--address-offset",
            "10",
            "-o",
            self.output_final,
        ]
        with patch.object(sys, "argv", test_args_run):
            main()

        self.assertTrue(os.path.exists(self.output_final))

        # Verify output Webdyn definition content has correct data and address offset
        with open(self.output_final, encoding="utf-8") as f:
            reader = csv.reader(f, delimiter=";")
            rows = list(reader)
            self.assertTrue(len(rows) > 1)
            # Header
            self.assertEqual(rows[0][2], "RunMfg")
            self.assertEqual(rows[0][3], "RunModel")

            # First register row
            # Info3 is the Modbus address in Webdyn format
            self.assertEqual(rows[1][5], "Power")

            # The modbus address should be original address + offset
            # 40001 (mapped address) + 10 = 40011
            self.assertEqual(rows[1][2], "40011")


if __name__ == "__main__":
    unittest.main()
