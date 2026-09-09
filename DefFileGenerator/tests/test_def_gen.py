import logging
import os
import unittest
from unittest.mock import patch

from DefFileGenerator.def_gen import Generator


class TestGenerator(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        # Suppress logging during tests unless checking for logs
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)

    def test_calculate_coefficients(self):
        # Default conditions
        self.assertEqual(
            self.generator._calculate_coefficients("", "", ""), ("1.000000", "0.000000")
        )
        self.assertEqual(
            self.generator._calculate_coefficients(None, None, None), ("1.000000", "0.000000")
        )

        # Valid standard inputs
        self.assertEqual(
            self.generator._calculate_coefficients("2", "10", "0"), ("2.000000", "10.000000")
        )

        # With scale factor
        self.assertEqual(
            self.generator._calculate_coefficients("1", "0", "2"), ("100.000000", "0.000000")
        )
        self.assertEqual(
            self.generator._calculate_coefficients("1.5", "0", "1"), ("15.000000", "0.000000")
        )
        self.assertEqual(
            self.generator._calculate_coefficients("1", "5", "-1"), ("0.100000", "5.000000")
        )

        # With scale factor formatted as float string
        self.assertEqual(
            self.generator._calculate_coefficients("2", "0", "2.0"), ("200.000000", "0.000000")
        )
        self.assertEqual(
            self.generator._calculate_coefficients("2", "0", "-2.0"), ("0.020000", "0.000000")
        )

        # Invalid scale factors fallback to 0
        self.assertEqual(
            self.generator._calculate_coefficients("2", "1", "invalid"), ("2.000000", "1.000000")
        )

        # Precision/Format up to 6 decimal places
        self.assertEqual(
            self.generator._calculate_coefficients("0.1234567", "0.9876543", "0"),
            ("0.123457", "0.987654"),
        )
        self.assertEqual(
            self.generator._calculate_coefficients("1", "0", "-3"), ("0.001000", "0.000000")
        )

    def test_intelligent_defaulting(self):
        rows = [
            {"Name": "Holding", "RegisterType": "Holding", "Address": "100", "Type": "U16"},
            {"Name": "Input", "RegisterType": "Input", "Address": "101", "Type": "U16"},
            {"Name": "Coil", "RegisterType": "Coil", "Address": "1", "Type": "U16"},
            {"Name": "Discrete", "RegisterType": "Discrete", "Address": "2", "Type": "U16"},
        ]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]["Action"], "1")  # Holding -> Read/Write
        self.assertEqual(processed[1]["Action"], "4")  # Input -> Read Only
        self.assertEqual(processed[2]["Action"], "1")  # Coil -> Read/Write
        self.assertEqual(processed[3]["Action"], "4")  # Discrete -> Read Only

    def test_validate_address_range(self):
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))

        logging.disable(logging.NOTSET)
        with self.assertLogs(level="WARNING") as log:
            self.assertFalse(self.generator.validate_address("65536", "U16"))
            self.assertTrue(any("out of Modbus range" in m for m in log.output))
        logging.disable(logging.CRITICAL)

    def test_normalize_address_val(self):
        self.assertEqual(self.generator.normalize_address_val("0x10"), "16")
        self.assertEqual(self.generator.normalize_address_val("10h"), "16")
        self.assertEqual(self.generator.normalize_address_val("10"), "10")
        self.assertEqual(self.generator.normalize_address_val("A0"), "160")
        self.assertEqual(self.generator.normalize_address_val("1,234"), "1234")

    def test_validate_address_invalid(self):
        self.assertFalse(self.generator.validate_address("30001_10", "U16"))  # U16 expects int
        self.assertFalse(self.generator.validate_address("xyz", "U16"))  # Not hex

    def test_get_register_count(self):
        self.assertEqual(self.generator.get_register_count("U16", "30000"), 1)
        self.assertEqual(self.generator.get_register_count("U32", "30000"), 2)
        self.assertEqual(self.generator.get_register_count("U64", "30000"), 4)
        self.assertEqual(self.generator.get_register_count("MAC", "30000"), 3)
        self.assertEqual(self.generator.get_register_count("IPV6", "30000"), 8)
        self.assertEqual(self.generator.get_register_count("STRING", "30000_10"), 5)  # ceil(10/2)
        self.assertEqual(self.generator.get_register_count("STRING", "30000_11"), 6)  # ceil(11/2)

    def test_apply_address_offset(self):
        # Empty input
        self.assertEqual(Generator.apply_address_offset("", 10), "")
        self.assertEqual(Generator.apply_address_offset(None, 10), "")

        # Simple address
        self.assertEqual(Generator.apply_address_offset("30000", 10), "30010")

        # Compound address
        self.assertEqual(Generator.apply_address_offset("40000_20", -5), "39995_20")

        # Hex addresses (normalized internally to decimal before offset)
        self.assertEqual(Generator.apply_address_offset("0x10", 5), "21")  # 16 + 5 = 21
        self.assertEqual(Generator.apply_address_offset("0x10_20", 5), "21_20")

        # Value Error handling (caught silently)
        self.assertEqual(Generator.apply_address_offset("invalid", 10), "invalid")
        self.assertEqual(Generator.apply_address_offset("invalid_20", 10), "invalid_20")

        # Negative address warning (with log assertions)
        logging.disable(logging.NOTSET)

        with self.assertLogs(level="WARNING") as log1:
            res1 = Generator.apply_address_offset("10", -15)
            self.assertEqual(res1, "-5")
            self.assertTrue(
                any("Address offset -15 results in negative address -5" in m for m in log1.output)
            )
            self.assertFalse(any("for '" in m for m in log1.output))
            self.assertFalse(any("Line" in m for m in log1.output))

        with self.assertLogs(level="WARNING") as log2:
            res2 = Generator.apply_address_offset("10", -15, line_num=42, name="TestReg")
            self.assertEqual(res2, "-5")
            self.assertTrue(
                any(
                    "Line 42: Address offset -15 results in negative address -5 for 'TestReg'" in m
                    for m in log2.output
                )
            )

        logging.disable(logging.CRITICAL)

    def test_process_rows_basic(self):
        rows = [
            {
                "Name": "Test Var",
                "Tag": "test_tag",
                "RegisterType": "Holding Register",
                "Address": "30000",
                "Type": "U16",
                "Factor": "1",
                "Offset": "0",
                "Unit": "V",
                "Action": "4",
                "ScaleFactor": "0",
            }
        ]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(len(processed), 1)
        self.assertEqual(processed[0]["Info1"], "3")
        self.assertEqual(processed[0]["Info3"], "U16")
        self.assertEqual(processed[0]["CoefA"], "1.000000")

    def test_automatic_tag_generation(self):
        rows = [
            {
                "Name": "Test Variable",
                "Tag": "",
                "RegisterType": "3",
                "Address": "100",
                "Type": "U16",
            },
            {
                "Name": "Test Variable",
                "Tag": "",
                "RegisterType": "3",
                "Address": "101",
                "Type": "U16",
            },
        ]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]["Tag"], "test_variable")
        self.assertEqual(processed[1]["Tag"], "test_variable_1")

    def test_action_normalization(self):
        rows = [
            {
                "Name": "Var1",
                "Tag": "t1",
                "RegisterType": "3",
                "Address": "100",
                "Type": "U16",
                "Action": "R",
                "Factor": "",
                "Offset": "",
                "Unit": "",
                "ScaleFactor": "",
            },
            {
                "Name": "Var2",
                "Tag": "t2",
                "RegisterType": "3",
                "Address": "101",
                "Type": "U16",
                "Action": "RW",
                "Factor": "",
                "Offset": "",
                "Unit": "",
                "ScaleFactor": "",
            },
            {
                "Name": "Var3",
                "Tag": "t3",
                "RegisterType": "3",
                "Address": "102",
                "Type": "U16",
                "Action": "write",
                "Factor": "",
                "Offset": "",
                "Unit": "",
                "ScaleFactor": "",
            },
        ]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]["Action"], "4")  # R -> 4
        self.assertEqual(processed[1]["Action"], "1")  # RW -> 1
        self.assertEqual(processed[2]["Action"], "1")  # write -> 1

    def test_generate_template_modes(self):
        import os

        from DefFileGenerator.def_gen import generate_template

        # Test input mode
        out_input = "template_input.csv"
        generate_template(out_input, mode="input")
        with open(out_input) as f:
            content = f.read()
            self.assertIn("Name,Tag,RegisterType", content)
        os.remove(out_input)

        # Test definition mode
        out_def = "template_def.csv"
        generate_template(out_def, mode="definition")
        with open(out_def) as f:
            content = f.read()
            self.assertIn("modbusRTU;Inverter", content)
        os.remove(out_def)

    def test_sanitize_csv_field(self):
        # Numeric values should be preserved without prepended apostrophe
        self.assertEqual(self.generator.sanitize_csv_field("-40.0"), "-40.0")
        self.assertEqual(self.generator.sanitize_csv_field("+1.23"), "+1.23")
        self.assertEqual(self.generator.sanitize_csv_field("-100"), "-100")
        self.assertEqual(self.generator.sanitize_csv_field("123"), "123")
        self.assertEqual(self.generator.sanitize_csv_field(None), "")

        # Formulas and injection attempts should have apostrophe prepended
        self.assertEqual(self.generator.sanitize_csv_field("=SUM(A1:A5)"), "'=SUM(A1:A5)")
        self.assertEqual(self.generator.sanitize_csv_field("+1+1"), "'+1+1")
        self.assertEqual(self.generator.sanitize_csv_field("-some_var"), "'-some_var")
        self.assertEqual(self.generator.sanitize_csv_field("@SUM"), "'@SUM")
        self.assertEqual(self.generator.sanitize_csv_field("-10.5"), "-10.5")
        self.assertEqual(self.generator.sanitize_csv_field("+25"), "+25")
        self.assertEqual(self.generator.sanitize_csv_field("-text"), "'-text")
        # Padded CSV injection test cases
        self.assertEqual(self.generator.sanitize_csv_field("   =1+1"), "'   =1+1")
        self.assertEqual(self.generator.sanitize_csv_field("  @SUM"), "'  @SUM")
        self.assertEqual(self.generator.sanitize_csv_field("\t+cmd"), "'\t+cmd")
        self.assertEqual(self.generator.sanitize_csv_field("   -text"), "'   -text")
        self.assertEqual(self.generator.sanitize_csv_field("   -10.5"), "   -10.5")
        self.assertEqual(self.generator.sanitize_csv_field("  +25"), "  +25")

    def test_write_output_csv_os_error_on_open(self):
        logging.disable(logging.NOTSET)
        with patch("builtins.open", side_effect=OSError("Permission denied")):
            with self.assertLogs(level="ERROR") as log:
                Generator.write_output_csv("invalid_path.csv", [], "Mfg", "Model")
                self.assertTrue(
                    any("Error writing output CSV: Permission denied" in m for m in log.output)
                )

    def test_write_output_csv_error_during_write_and_cleanup(self):
        import csv
        from unittest.mock import MagicMock

        file_obj = MagicMock()

        logging.disable(logging.NOTSET)
        with patch("builtins.open", return_value=file_obj):
            with patch("csv.writer") as mock_writer_cls:
                mock_writer = MagicMock()
                mock_writer.writerow.side_effect = csv.Error("CSV write error")
                mock_writer_cls.return_value = mock_writer

                with self.assertLogs(level="ERROR") as log:
                    Generator.write_output_csv("dummy.csv", [], "Mfg", "Model")
                    self.assertTrue(
                        any("Error writing output CSV: CSV write error" in m for m in log.output)
                    )
                file_obj.close.assert_called_once()


if __name__ == "__main__":
    unittest.main()


class TestGeneratorUncoveredEdgeCases(unittest.TestCase):
    def test_validate_csv_nonexistent_file(self):
        generator = Generator()
        logging.disable(logging.NOTSET)
        with self.assertLogs(level="ERROR") as log:
            res = generator.validate_csv("nonexistent_def_file.csv")
            self.assertFalse(res)
            self.assertTrue(
                any("File not found: nonexistent_def_file.csv" in m for m in log.output)
            )
        logging.disable(logging.CRITICAL)

    def test_validate_csv_malformed_lines(self):
        import tempfile

        generator = Generator()
        with tempfile.NamedTemporaryFile("w+", delete=False, suffix=".csv") as f:
            f.write("modbusRTU;Inverter\n")
            f.write("1;3;100;U16;1;Name;tag1;1.000000;0.000000;V;4\n")
            temp_path = f.name

        try:
            logging.disable(logging.NOTSET)
            res = generator.validate_csv(temp_path, strict=True)
            self.assertTrue(res)
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)
            logging.disable(logging.CRITICAL)
