import unittest
import os
import csv
import logging
import tempfile
from DefFileGenerator.def_gen import generate_template

class TestTemplate(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_generate_template(self):
        path = os.path.join(self.temp_dir.name, "template.csv")
        generate_template(path)

        self.assertTrue(os.path.exists(path))
        with open(path, "r", newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            rows = list(reader)
            self.assertGreater(len(rows), 0)
            self.assertIn("Name", rows[0])
            self.assertIn("Address", rows[0])
            self.assertIn("Type", rows[0])

if __name__ == "__main__":
    unittest.main()
