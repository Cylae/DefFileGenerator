import unittest
import os
import csv
import sys
import logging
from unittest.mock import patch
from DefFileGenerator.main import main

class TestValidateCommand(unittest.TestCase):
    def setUp(self):
        self.def_file = "test_validate.csv"
        # Standard Webdyn definition format: 11 columns, semicolon delimited
        # Header: protocol, category, manufacturer, model, forced_write, ...
        with open(self.def_file, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'TestMfg', 'TestModel', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U16', '', 'Var1', 'tag1', '1.0', '0.0', 'V', '4'])

    def tearDown(self):
        if os.path.exists(self.def_file):
            os.remove(self.def_file)

    def test_validate_success(self):
        test_args = ["main.py", "validate", self.def_file]
        with patch.object(sys, 'argv', test_args):
            with self.assertLogs(level='INFO') as log:
                try:
                    main()
                except SystemExit as e:
                    self.assertEqual(e.code, 0)
                self.assertTrue(any("Validation successful" in m for m in log.output))

    def test_validate_duplicate_tag(self):
        with open(self.def_file, 'a', encoding='utf-8', newline='') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['2', '3', '101', 'U16', '', 'Var2', 'tag1', '1.0', '0.0', 'V', '4'])

        test_args = ["main.py", "validate", self.def_file]
        with patch.object(sys, 'argv', test_args):
            with self.assertLogs(level='ERROR') as log:
                with self.assertRaises(SystemExit) as cm:
                    main()
                self.assertEqual(cm.exception.code, 1)
                self.assertTrue(any("Fatal Duplicate Tag 'tag1'" in m for m in log.output))

    def test_validate_address_overlap(self):
        # 100 for U16 is 1 register. 100 for U32 is 2 registers (100, 101).
        with open(self.def_file, 'a', encoding='utf-8', newline='') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['2', '3', '100', 'U32', '', 'Var2', 'tag2', '1.0', '0.0', 'V', '4'])

        test_args = ["main.py", "validate", self.def_file]
        with patch.object(sys, 'argv', test_args):
            with self.assertLogs(level='WARNING') as log:
                main()
                self.assertTrue(any("Address overlap detected" in m for m in log.output))

if __name__ == '__main__':
    unittest.main()
