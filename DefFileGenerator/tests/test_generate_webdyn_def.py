import unittest
import sys
import os
import logging
from unittest.mock import patch

class TestGenerateWebdynDef(unittest.TestCase):
    def setUp(self):
        # Suppress logging to keep output clean, but allow test log assertion if needed
        self.old_log_level = logging.getLogger().getEffectiveLevel()
        logging.disable(logging.CRITICAL)

        self.input_csv = "test_gen_wrapper_input.csv"
        self.output_csv = "test_gen_wrapper_output.csv"

        with open(self.input_csv, 'w', encoding='utf-8') as f:
            f.write("Register,Name,Data Type,Unit,Scale,Access\n")
            f.write("30001,Power,uint16,W,1,R\n")
            f.write("30002,Voltage,uint16,V,0.1,R\n")

    def tearDown(self):
        # Restore logging level
        logging.disable(self.old_log_level)

        # Cleanup temporary files
        for path in [self.input_csv, self.output_csv, "sample_register_map.csv", "sample_output_definition.csv", "test_out_offset.csv"]:
            if os.path.exists(path):
                os.remove(path)

    def test_generate_webdyn_definition_success(self):
        from generate_webdyn_def import generate_webdyn_definition
        success = generate_webdyn_definition(
            input_file=self.input_csv,
            output_file=self.output_csv,
            manufacturer="TestMfg",
            model="TestModel"
        )
        self.assertTrue(success)
        self.assertTrue(os.path.exists(self.output_csv))

        # Check content of the generated file
        with open(self.output_csv, 'r', encoding='utf-8-sig') as f:
            lines = f.readlines()
        self.assertTrue(len(lines) >= 3)
        self.assertTrue("TestMfg" in lines[0])
        self.assertTrue("TestModel" in lines[0])

    def test_generate_webdyn_definition_nonexistent_input(self):
        from generate_webdyn_def import generate_webdyn_definition
        success = generate_webdyn_definition(
            input_file="nonexistent_file.csv",
            output_file=self.output_csv,
            manufacturer="TestMfg",
            model="TestModel"
        )
        self.assertFalse(success)
        self.assertFalse(os.path.exists(self.output_csv))

    def test_generate_webdyn_definition_unsupported_format(self):
        from generate_webdyn_def import generate_webdyn_definition
        bad_format_file = "test_bad_ext.txt"
        with open(bad_format_file, 'w') as f:
            f.write("some data")

        try:
            success = generate_webdyn_definition(
                input_file=bad_format_file,
                output_file=self.output_csv,
                manufacturer="TestMfg",
                model="TestModel"
            )
            self.assertFalse(success)
        finally:
            if os.path.exists(bad_format_file):
                os.remove(bad_format_file)

    def test_generate_webdyn_definition_empty_input(self):
        from generate_webdyn_def import generate_webdyn_definition
        empty_file = "test_empty_input.csv"
        with open(empty_file, 'w') as f:
            f.write("")

        try:
            success = generate_webdyn_definition(
                input_file=empty_file,
                output_file=self.output_csv,
                manufacturer="TestMfg",
                model="TestModel"
            )
            self.assertFalse(success)
        finally:
            if os.path.exists(empty_file):
                os.remove(empty_file)

    def test_generate_webdyn_definition_address_offset(self):
        from generate_webdyn_def import generate_webdyn_definition
        success = generate_webdyn_definition(
            input_file=self.input_csv,
            output_file="test_out_offset.csv",
            manufacturer="TestMfg",
            model="TestModel",
            address_offset=10
        )
        self.assertTrue(success)
        self.assertTrue(os.path.exists("test_out_offset.csv"))

        # Verify that addresses are shifted (30001 -> 30011, 30002 -> 30012)
        with open("test_out_offset.csv", 'r', encoding='utf-8-sig') as f:
            lines = f.readlines()
        self.assertTrue(any(";30011;" in line for line in lines))
        self.assertTrue(any(";30012;" in line for line in lines))

    def test_generate_webdyn_definition_validation_fail_strict(self):
        # Create an input with overlapping addresses to cause validation failure in strict mode
        overlap_csv = "test_overlap.csv"
        with open(overlap_csv, 'w', encoding='utf-8') as f:
            f.write("Register,Name,Data Type\n")
            f.write("40001,ActivePower,uint16\n")
            f.write("40001,ReactivePower,uint16\n")

        from generate_webdyn_def import generate_webdyn_definition
        try:
            success = generate_webdyn_definition(
                input_file=overlap_csv,
                output_file=self.output_csv,
                manufacturer="TestMfg",
                model="TestModel",
                strict_validation=True
            )
            # Should fail due to address overlap
            self.assertFalse(success)
        finally:
            if os.path.exists(overlap_csv):
                os.remove(overlap_csv)

    def test_main_demo_mode(self):
        # When fewer than 5 arguments are provided, main should run the demo mode
        from generate_webdyn_def import main
        test_args = ['generate_webdyn_def.py']
        with patch.object(sys, 'argv', test_args):
            with patch('sys.exit') as mock_exit:
                main()
                mock_exit.assert_called_with(0)
                self.assertTrue(os.path.exists("sample_register_map.csv"))
                self.assertTrue(os.path.exists("sample_output_definition.csv"))

    def test_main_with_arguments(self):
        from generate_webdyn_def import main
        test_args = ['generate_webdyn_def.py', self.input_csv, self.output_csv, 'TestMfg', 'TestModel']
        with patch.object(sys, 'argv', test_args):
            with patch('sys.exit') as mock_exit:
                main()
                mock_exit.assert_called_with(0)
                self.assertTrue(os.path.exists(self.output_csv))

    def test_main_with_arguments_and_optionals(self):
        from generate_webdyn_def import main
        test_args = [
            'generate_webdyn_def.py',
            self.input_csv,
            self.output_csv,
            'TestMfg',
            'TestModel',
            'modbusRTU',
            'Sensor'
        ]
        with patch.object(sys, 'argv', test_args):
            with patch('sys.exit') as mock_exit:
                main()
                mock_exit.assert_called_with(0)
                self.assertTrue(os.path.exists(self.output_csv))

                # Check category in header row
                with open(self.output_csv, 'r', encoding='utf-8-sig') as f:
                    header = f.readline()
                self.assertTrue(";Sensor;" in header)

    def test_module_execution(self):
        import runpy
        test_args = ['generate_webdyn_def.py', self.input_csv, self.output_csv, 'TestMfg', 'TestModel']
        with patch.object(sys, 'argv', test_args):
            with patch('sys.exit') as mock_exit:
                runpy.run_module('generate_webdyn_def', run_name='__main__')
                mock_exit.assert_called_with(0)
                self.assertTrue(os.path.exists(self.output_csv))

if __name__ == "__main__":
    unittest.main()