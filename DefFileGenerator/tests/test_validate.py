import unittest
import os
import csv
import logging
from DefFileGenerator.def_gen import Generator, peek_generator
from DefFileGenerator.main import main as cli_main
from io import StringIO
from unittest.mock import patch

class TestValidateCommand(unittest.TestCase):
    def setUp(self):
        self.test_file = "test_validate_cli.csv"
        # Disable logging for tests to keep output clean
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
        logging.disable(logging.NOTSET)

    def test_validate_success(self):
        with open(self.test_file, 'w', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Test', 'Model', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])

        with patch('sys.argv', ['main.py', 'validate', self.test_file]):
            with patch('sys.stdout', new=StringIO()):
                try:
                    cli_main()
                except SystemExit as e:
                    self.assertEqual(e.code, 0)

    def test_validate_failure(self):
        with open(self.test_file, 'w', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'Test', 'Model', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])
            writer.writerow(['2', '3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4']) # Duplicate tag

        with patch('sys.argv', ['main.py', 'validate', self.test_file]):
            with patch('sys.stderr', new=StringIO()):
                with self.assertRaises(SystemExit) as cm:
                    cli_main()
                self.assertEqual(cm.exception.code, 1)

if __name__ == "__main__":
    unittest.main()
