import unittest
import os
import csv
import logging
import io
import sys
from unittest.mock import patch
from DefFileGenerator.def_gen import Generator

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.test_file = 'test_validate.csv'
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
        logging.disable(logging.NOTSET)

    def create_csv(self, rows, header=None):
        with open(self.test_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            if header is None:
                writer.writerow(['modbusRTU', 'Inverter', 'Test', 'Model', '', '', '', '', '', '', ''])
            else:
                writer.writerow(header)
            for row in rows:
                writer.writerow(row)

    def test_valid_file(self):
        self.create_csv([
            ['1', '3', '40001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['2', '3', '40002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']
        ])
        self.assertTrue(Generator.validate_csv(self.test_file))

    def test_duplicate_tag(self):
        self.create_csv([
            ['1', '3', '40001', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['2', '3', '40002', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'A', '4']
        ])
        self.assertFalse(Generator.validate_csv(self.test_file))

    def test_address_overlap(self):
        # U32 takes 2 registers. 40001 and 40002.
        self.create_csv([
            ['1', '3', '40001', 'U32', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'],
            ['2', '3', '40002', 'U16', '', 'Var2', 'tag2', '1.0', '0.0', 'A', '4']
        ])
        # Overlap is a warning, so validate_csv still returns True but logs it.
        self.assertTrue(Generator.validate_csv(self.test_file))

    def test_invalid_address(self):
        self.create_csv([
            ['1', '3', 'not_an_addr', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        self.assertFalse(Generator.validate_csv(self.test_file))

    def test_address_out_of_range(self):
        self.create_csv([
            ['1', '3', '70000', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4']
        ])
        # Warning only, so returns True
        self.assertTrue(Generator.validate_csv(self.test_file))

    def test_empty_file(self):
        with open(self.test_file, 'w') as f: pass
        self.assertFalse(Generator.validate_csv(self.test_file))

    def test_invalid_header(self):
        self.create_csv([], header=['short', 'header'])
        self.assertFalse(Generator.validate_csv(self.test_file))

    def test_file_not_found(self):
        self.assertFalse(Generator.validate_csv('nonexistent.csv'))

    def test_template_generation(self):
        from doc_to_webdyn import main
        test_args = ["doc_to_webdyn.py", "--template", "-o", "template.csv"]
        with patch.object(sys, 'argv', test_args):
            main()
        self.assertTrue(os.path.exists("template.csv"))
        os.remove("template.csv")

    def test_template_stdout(self):
        from doc_to_webdyn import main
        test_args = ["doc_to_webdyn.py", "--template"]
        with patch.object(sys, 'argv', test_args):
            captured_output = io.StringIO()
            sys.stdout = captured_output
            try:
                main()
            finally:
                sys.stdout = sys.__stdout__
            self.assertIn("Example Variable", captured_output.getvalue())

if __name__ == '__main__':
    unittest.main()
