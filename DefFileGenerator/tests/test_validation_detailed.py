"""Tests for detailed CSV validation and diagnostics reporting."""

import tempfile
import unittest

from DefFileGenerator.def_gen import Generator, ValidationReport


class TestValidationDetailed(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_validate_csv_detailed_valid_file(self):
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".csv", delete=False, newline="", encoding="utf-8"
        ) as f:
            f.write("modbusRTU;Inverter;Huawei;SUN2000;;;;;;;\n")
            f.write("1;3;40001;U16;;Active Power;active_power;1.0;0.0;W;4\n")
            f.write("2;3;40002;U16;;Grid Voltage;grid_voltage;0.1;0.0;V;4\n")
            filepath = f.name

        report = self.generator.validate_csv_detailed(filepath, strict=True)
        self.assertIsInstance(report, ValidationReport)
        self.assertTrue(report.is_valid)
        self.assertEqual(report.register_count, 2)
        self.assertEqual(report.stats["errors"], 0)
        self.assertEqual(len(report.issues), 0)

    def test_validate_csv_detailed_duplicate_tag(self):
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".csv", delete=False, newline="", encoding="utf-8"
        ) as f:
            f.write("modbusRTU;Inverter;Huawei;SUN2000;;;;;;;\n")
            f.write("1;3;40001;U16;;Power 1;power_tag;1.0;0.0;W;4\n")
            f.write("2;3;40002;U16;;Power 2;power_tag;1.0;0.0;W;4\n")
            filepath = f.name

        report = self.generator.validate_csv_detailed(filepath, strict=True)
        self.assertFalse(report.is_valid)
        self.assertEqual(report.stats["errors"], 1)
        issue = report.issues[0]
        self.assertEqual(issue.line, 3)
        self.assertEqual(issue.code, "DUPLICATE_TAG")
        self.assertEqual(issue.field, "Tag")
        self.assertEqual(issue.severity, "ERROR")

    def test_validate_csv_detailed_invalid_type(self):
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".csv", delete=False, newline="", encoding="utf-8"
        ) as f:
            f.write("modbusRTU;Inverter;Huawei;SUN2000;;;;;;;\n")
            f.write("1;3;40001;INVALID_TYPE_XYZ;;Power;pwr;1.0;0.0;W;4\n")
            filepath = f.name

        report = self.generator.validate_csv_detailed(filepath, strict=True)
        self.assertFalse(report.is_valid)
        self.assertTrue(any(i.code == "INVALID_TYPE" for i in report.issues))

    def test_validate_csv_detailed_insufficient_columns(self):
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".csv", delete=False, newline="", encoding="utf-8"
        ) as f:
            f.write("modbusRTU;Inverter;Huawei;SUN2000;;;;;;;\n")
            f.write("1;3;40001;U16\n")  # Only 4 columns instead of 11
            filepath = f.name

        report = self.generator.validate_csv_detailed(filepath, strict=True)
        self.assertFalse(report.is_valid)
        self.assertTrue(any(i.code == "INSUFFICIENT_COLUMNS" for i in report.issues))

    def test_validate_csv_detailed_address_overlap(self):
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".csv", delete=False, newline="", encoding="utf-8"
        ) as f:
            f.write("modbusRTU;Inverter;Huawei;SUN2000;;;;;;;\n")
            f.write(
                "1;3;40001;U32_WB;;Power 32bit;pwr_32;1.0;0.0;W;4\n"
            )  # consumes 40001 and 40002
            f.write("2;3;40002;U16;;Voltage 16bit;volt_16;1.0;0.0;V;4\n")  # overlaps with 40002
            filepath = f.name

        report_lenient = self.generator.validate_csv_detailed(
            filepath, strict=False, strict_overlap=False
        )
        self.assertTrue(report_lenient.is_valid)
        self.assertTrue(
            any(
                i.code == "ADDRESS_OVERLAP" and i.severity == "WARNING"
                for i in report_lenient.issues
            )
        )

        report_strict = self.generator.validate_csv_detailed(
            filepath, strict=True, strict_overlap=True
        )
        self.assertFalse(report_strict.is_valid)
        self.assertTrue(any(i.code == "ADDRESS_OVERLAP" for i in report_strict.issues))

    def test_validate_csv_detailed_missing_file(self):
        report = self.generator.validate_csv_detailed("non_existent_file_9999.csv")
        self.assertFalse(report.is_valid)
        self.assertEqual(report.issues[0].code, "FILE_NOT_FOUND")


if __name__ == "__main__":
    unittest.main()
