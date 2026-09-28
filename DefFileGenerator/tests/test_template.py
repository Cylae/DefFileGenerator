import csv
import os
import tempfile
import unittest

from DefFileGenerator.def_gen import Generator, generate_template


class TestTemplate(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_generate_template(self):
        path = os.path.join(self.temp_dir.name, "template.csv")
        generate_template(path)

        self.assertTrue(os.path.exists(path))
        with open(path, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            rows = list(reader)
            self.assertGreater(len(rows), 0)
            self.assertIn("Name", rows[0])
            self.assertIn("Address", rows[0])
            self.assertIn("Type", rows[0])

    def test_definition_template_is_a_valid_definition(self):
        path = os.path.join(self.temp_dir.name, "definition-template.csv")

        generate_template(path, mode="definition")

        report = Generator().validate_csv_detailed(path, strict=True)
        self.assertTrue(report.is_valid, report.issues)
        self.assertEqual(report.register_count, 2)


if __name__ == "__main__":
    unittest.main()
