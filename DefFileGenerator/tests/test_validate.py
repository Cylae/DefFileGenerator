import unittest
import sys
import os
import io
from unittest.mock import patch

class TestValidate(unittest.TestCase):
    def setUp(self):
        self.valid_csv = "test_valid.csv"
        with open(self.valid_csv, 'w', encoding='utf-8') as f:
            f.write("modbusRTU;Inverter;TestMfg;TestModel;;;;;;;\n")
            f.write("1;3;100;U16;;Test;test_tag;1.0;0.0;V;4\n")

        self.invalid_csv = "test_invalid.csv"
        with open(self.invalid_csv, 'w', encoding='utf-8') as f:
            f.write("modbusRTU;Inverter;TestMfg;TestModel;;;;;;;\n")
            f.write("1;3;100;U16;;Test;test_tag;1.0;0.0;V;4\n")
            f.write("2;3;100;U16;;Test2;test_tag;1.0;0.0;V;4\n") # Duplicate tag 'test_tag'

    def tearDown(self):
        if os.path.exists(self.valid_csv):
            os.remove(self.valid_csv)
        if os.path.exists(self.invalid_csv):
            os.remove(self.invalid_csv)

    def test_validate_success(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "validate", self.valid_csv]
        with patch.object(sys, 'argv', test_args):
            with self.assertLogs(level='INFO') as log:
                main()
                self.assertTrue(any("Validation successful" in m for m in log.output))

    def test_validate_fail(self):
        from DefFileGenerator.main import main
        test_args = ["main.py", "validate", self.invalid_csv]
        with patch.object(sys, 'argv', test_args):
            with self.assertLogs(level='ERROR') as log:
                with self.assertRaises(SystemExit) as cm:
                    main()
                self.assertEqual(cm.exception.code, 1)
                self.assertTrue(any("Duplicate Tag" in m for m in log.output))

if __name__ == "__main__":
    unittest.main()
