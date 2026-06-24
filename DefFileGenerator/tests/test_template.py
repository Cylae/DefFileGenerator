import unittest
import os
import csv
import sys
import io
from DefFileGenerator.main import main

class TestTemplate(unittest.TestCase):
    def tearDown(self):
        if os.path.exists("test_template.csv"):
            os.remove("test_template.csv")

    def test_generate_template_csv(self):
        test_args = ["main.py", "generate", "--template", "-o", "test_template.csv"]
        with unittest.mock.patch.object(sys, 'argv', test_args):
            main()
            self.assertTrue(os.path.exists("test_template.csv"))
            with open("test_template.csv", 'r') as f:
                reader = csv.reader(f)
                header = next(reader)
                self.assertEqual(header[0], "Name")
                row = next(reader)
                self.assertEqual(row[0], "Example Variable")

    def test_run_template_def(self):
        test_args = ["main.py", "run", "--template", "-o", "test_template.csv"]
        with unittest.mock.patch.object(sys, 'argv', test_args):
            main()
            self.assertTrue(os.path.exists("test_template.csv"))
            with open("test_template.csv", 'r') as f:
                reader = csv.reader(f, delimiter=';')
                header = next(reader)
                self.assertEqual(header[0], "modbusRTU")
                row = next(reader)
                self.assertEqual(row[1], "3") # Holding register

if __name__ == "__main__":
    unittest.main()
