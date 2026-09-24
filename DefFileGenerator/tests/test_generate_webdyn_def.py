import logging
import os
import sys
import unittest
from unittest.mock import patch


class TestGenerateWebdynDef(unittest.TestCase):
    def setUp(self):
        # Suppress logging to keep output clean, but allow test log assertion if needed
        self.old_log_level = logging.getLogger().getEffectiveLevel()
        logging.disable(logging.CRITICAL)

        self.input_csv = "test_gen_wrapper_input.csv"
        self.output_csv = "test_gen_wrapper_output.csv"

        with open(self.input_csv, "w", encoding="utf-8") as f:
            f.write("Register,Name,Data Type,Unit,Scale,Access\n")
            f.write("30001,Power,uint16,W,1,R\n")
            f.write("30002,Voltage,uint16,V,0.1,R\n")

    def tearDown(self):
        # Restore logging level
        logging.disable(self.old_log_level)

        # Cleanup temporary files
        for path in [
            self.input_csv,
            self.output_csv,
            "sample_register_map.csv",
            "sample_output_definition.csv",
            "test_out_offset.csv",
        ]:
            if os.path.exists(path):
                os.remove(path)

    def test_generate_webdyn_definition_success(self):
        from generate_webdyn_def import generate_webdyn_definition

        success = generate_webdyn_definition(
            input_file=self.input_csv,
            output_file=self.output_csv,
            manufacturer="TestMfg",
            model="TestModel",
        )
        self.assertTrue(success)
        self.assertTrue(os.path.exists(self.output_csv))

        # Check content of the generated file
        with open(self.output_csv, encoding="utf-8-sig") as f:
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
            model="TestModel",
        )
        self.assertFalse(success)
        self.assertFalse(os.path.exists(self.output_csv))

    def test_generate_webdyn_definition_unsupported_format(self):
        from generate_webdyn_def import generate_webdyn_definition

        bad_format_file = "test_bad_ext.txt"
        with open(bad_format_file, "w") as f:
            f.write("some data")

        try:
            success = generate_webdyn_definition(
                input_file=bad_format_file,
                output_file=self.output_csv,
                manufacturer="TestMfg",
                model="TestModel",
            )
            self.assertFalse(success)
        finally:
            if os.path.exists(bad_format_file):
                os.remove(bad_format_file)

    def test_generate_webdyn_definition_no_raw_data_extracted(self):
        from generate_webdyn_def import generate_webdyn_definition

        # Mock extract_from_csv to yield empty generator
        with patch(
            "DefFileGenerator.extractor.Extractor.extract_from_csv", return_value=(x for x in [])
        ):
            success = generate_webdyn_definition(
                input_file=self.input_csv,
                output_file=self.output_csv,
                manufacturer="TestMfg",
                model="TestModel",
            )
            self.assertFalse(success)

    def test_generate_webdyn_definition_unknown_extension(self):
        from generate_webdyn_def import generate_webdyn_definition

        # Create dummy file with .unknown extension
        unk_file = "test_unk.unknown"
        with open(unk_file, "w") as f:
            f.write("content")
        try:
            success = generate_webdyn_definition(
                input_file=unk_file,
                output_file=self.output_csv,
                manufacturer="TestMfg",
                model="TestModel",
            )
            self.assertFalse(success)
        finally:
            if os.path.exists(unk_file):
                os.remove(unk_file)

    def test_generate_webdyn_definition_empty_input(self):
        from generate_webdyn_def import generate_webdyn_definition

        empty_file = "test_empty_input.csv"
        with open(empty_file, "w") as f:
            f.write("")

        try:
            success = generate_webdyn_definition(
                input_file=empty_file,
                output_file=self.output_csv,
                manufacturer="TestMfg",
                model="TestModel",
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
            address_offset=10,
        )
        self.assertTrue(success)
        self.assertTrue(os.path.exists("test_out_offset.csv"))

        # Verify that addresses are shifted (30001 -> 30011, 30002 -> 30012)
        with open("test_out_offset.csv", encoding="utf-8-sig") as f:
            lines = f.readlines()
        self.assertTrue(any(";30011;" in line for line in lines))
        self.assertTrue(any(";30012;" in line for line in lines))

    def test_generate_webdyn_definition_validation_fail_strict(self):
        # Create an input with overlapping addresses to cause validation failure in strict mode
        overlap_csv = "test_overlap.csv"
        with open(overlap_csv, "w", encoding="utf-8") as f:
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
                strict_validation=True,
            )
            # Should fail due to address overlap
            self.assertFalse(success)
        finally:
            if os.path.exists(overlap_csv):
                os.remove(overlap_csv)

    def test_generate_webdyn_definition_no_valid_registers_after_mapping(self):
        from generate_webdyn_def import generate_webdyn_definition

        # Mock map_and_clean to yield empty generator
        with patch(
            "DefFileGenerator.extractor.Extractor.map_and_clean", return_value=(x for x in [])
        ):
            success = generate_webdyn_definition(
                input_file=self.input_csv,
                output_file=self.output_csv,
                manufacturer="TestMfg",
                model="TestModel",
            )
            self.assertFalse(success)

    def test_generate_webdyn_definition_validation_return_false(self):
        from generate_webdyn_def import generate_webdyn_definition

        # Mock Generator.validate_csv to return False
        with patch("DefFileGenerator.def_gen.Generator.validate_csv", return_value=False):
            success = generate_webdyn_definition(
                input_file=self.input_csv,
                output_file=self.output_csv,
                manufacturer="TestMfg",
                model="TestModel",
            )
            self.assertFalse(success)

    def test_generate_webdyn_definition_missing_required_params(self):
        from generate_webdyn_def import generate_webdyn_definition

        # Calling with string input_file without output_file, manufacturer, or model
        success = generate_webdyn_definition(input_file=self.input_csv)
        self.assertFalse(success)

    def test_generate_webdyn_definition_with_webdyn_def_config(self):
        from DefFileGenerator.def_gen import WebdynDefConfig
        from generate_webdyn_def import generate_webdyn_definition

        cfg = WebdynDefConfig(
            input_file=self.input_csv,
            output_file=self.output_csv,
            manufacturer="ConfigMfg",
            model="ConfigModel",
        )
        success = generate_webdyn_definition(cfg)
        self.assertTrue(success)
        self.assertTrue(os.path.exists(self.output_csv))

    def test_generate_webdyn_definition_other_formats(self):
        from generate_webdyn_def import generate_webdyn_definition

        # Test excel, pdf, xml formats error handling / routing
        for ext in [".xlsx", ".pdf", ".xml"]:
            fname = f"test_dummy{ext}"
            with open(fname, "w") as f:
                f.write("dummy content")
            try:
                with patch(
                    "DefFileGenerator.extractor.Extractor.extract_from_excel",
                    side_effect=Exception("Excel err"),
                ):
                    with patch(
                        "DefFileGenerator.extractor.Extractor.extract_from_pdf",
                        side_effect=Exception("PDF err"),
                    ):
                        with patch(
                            "DefFileGenerator.extractor.Extractor.extract_from_xml",
                            side_effect=Exception("XML err"),
                        ):
                            success = generate_webdyn_definition(
                                input_file=fname,
                                output_file=self.output_csv,
                                manufacturer="Mfg",
                                model="Model",
                            )
                            self.assertFalse(success)
            finally:
                if os.path.exists(fname):
                    os.remove(fname)

    def test_generate_webdyn_definition_generator_exception(self):
        from generate_webdyn_def import generate_webdyn_definition

        with patch("generate_webdyn_def.run_generator", side_effect=RuntimeError("Gen failed")):
            success = generate_webdyn_definition(
                input_file=self.input_csv,
                output_file=self.output_csv,
                manufacturer="TestMfg",
                model="TestModel",
            )
            self.assertFalse(success)

    def test_generate_webdyn_definition_non_strict_validation(self):
        # Input with non-fatal warning (e.g. unknown tag or duplicate name warning if non-strict)
        warn_csv = "test_warn.csv"
        with open(warn_csv, "w", encoding="utf-8") as f:
            f.write("Register,Name,Data Type,Unit\n")
            f.write("30001,Power,uint16,W\n")

        from generate_webdyn_def import generate_webdyn_definition

        try:
            success = generate_webdyn_definition(
                input_file=warn_csv,
                output_file=self.output_csv,
                manufacturer="TestMfg",
                model="TestModel",
                strict_validation=False,
            )
            self.assertTrue(success)
        finally:
            if os.path.exists(warn_csv):
                os.remove(warn_csv)

    def test_main_demo_mode(self):
        # When fewer than 5 arguments are provided, main should run the demo mode
        from generate_webdyn_def import main

        test_args = ["generate_webdyn_def.py"]
        with patch.object(sys, "argv", test_args):
            with patch("sys.exit") as mock_exit:
                main()
                mock_exit.assert_called_with(0)
                self.assertTrue(os.path.exists("sample_register_map.csv"))
                self.assertTrue(os.path.exists("sample_output_definition.csv"))

    def test_main_with_arguments(self):
        from generate_webdyn_def import main

        test_args = [
            "generate_webdyn_def.py",
            self.input_csv,
            self.output_csv,
            "TestMfg",
            "TestModel",
        ]
        with patch.object(sys, "argv", test_args):
            with patch("sys.exit") as mock_exit:
                main()
                mock_exit.assert_called_with(0)
                self.assertTrue(os.path.exists(self.output_csv))

    def test_main_with_arguments_and_optionals(self):
        from generate_webdyn_def import main

        test_args = [
            "generate_webdyn_def.py",
            self.input_csv,
            self.output_csv,
            "TestMfg",
            "TestModel",
            "modbusRTU",
            "Sensor",
        ]
        with patch.object(sys, "argv", test_args):
            with patch("sys.exit") as mock_exit:
                main()
                mock_exit.assert_called_with(0)
                self.assertTrue(os.path.exists(self.output_csv))

                # Check category in header row
                with open(self.output_csv, encoding="utf-8-sig") as f:
                    header = f.readline()
                self.assertTrue(";Sensor;" in header)

    def test_module_execution(self):
        import runpy

        test_args = [
            "generate_webdyn_def.py",
            self.input_csv,
            self.output_csv,
            "TestMfg",
            "TestModel",
        ]
        with patch.object(sys, "argv", test_args):
            with patch("sys.exit") as mock_exit:
                runpy.run_module("generate_webdyn_def", run_name="__main__")
                mock_exit.assert_called_with(0)
                self.assertTrue(os.path.exists(self.output_csv))


if __name__ == "__main__":
    unittest.main()
