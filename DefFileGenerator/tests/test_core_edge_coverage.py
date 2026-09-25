import io
import os
import tempfile
import unittest

from DefFileGenerator.def_gen import (
    CSVHeaderConfig,
    Generator,
    GeneratorConfig,
    WebdynDefConfig,
    run_generator,
)
from DefFileGenerator.extractor import Extractor


class TestCoreEdgeCoverage(unittest.TestCase):
    def test_parse_numeric_fractions_and_invalid_inputs(self):
        self.assertEqual(Generator._parse_numeric(None), 0.0)
        self.assertEqual(Generator._parse_numeric(""), 0.0)
        self.assertEqual(Generator._parse_numeric("1/10"), 0.1)
        self.assertEqual(Generator._parse_numeric("1/0"), 0.0)
        self.assertEqual(Generator._parse_numeric("1/2/3"), 0.0)
        self.assertEqual(Generator._parse_numeric("invalid_string"), 0.0)

    def test_calculate_coefficients_extreme_exponent_and_non_finite(self):
        # Scale factor > 100 or < -100 resets scale to 0
        coef_a, coef_b = Generator._calculate_coefficients("1.5", "10.0", "150")
        self.assertEqual(coef_a, "1.500000")
        self.assertEqual(coef_b, "10.000000")

        coef_a, coef_b = Generator._calculate_coefficients("nan", "inf", "0")
        self.assertEqual(coef_a, "1.000000")
        self.assertEqual(coef_b, "0.000000")

    def test_apply_address_offset_empty_and_zero_offset(self):
        self.assertEqual(Generator.apply_address_offset("", 10), "")
        self.assertEqual(Generator.apply_address_offset(None, 10), "")
        self.assertEqual(Generator.apply_address_offset("40001", 0), "40001")
        self.assertEqual(Generator.apply_address_offset("40001_10", 0), "40001_10")

    def test_dataclass_post_init_normalization(self):
        web_cfg = WebdynDefConfig(
            input_file="in.csv", output_file="out.csv", manufacturer="M", model="M"
        )
        self.assertEqual(web_cfg.input_file, "in.csv")
        self.assertEqual(web_cfg.output_file, "out.csv")

        gen_cfg = GeneratorConfig(input_file="in.csv", output="out.csv")
        self.assertEqual(gen_cfg.input_file, "in.csv")
        self.assertEqual(gen_cfg.output, "out.csv")

    def test_write_output_csv_with_io_error(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            invalid_path = os.path.join(tmp_dir, "non_existent_subdir", "output.csv")
            # Attempting to write to non-existent directory triggers OSError in write_output_csv
            Generator.write_output_csv(
                invalid_path, [], CSVHeaderConfig(manufacturer="Mfg", model="Model")
            )
            self.assertFalse(os.path.exists(invalid_path))

    def test_write_output_csv_legacy_positional_string_config(self):
        buf = io.StringIO()
        rows = [
            {
                "Info1": "3",
                "Info2": "40001",
                "Info3": "U16",
                "Info4": "",
                "Name": "Volts",
                "Tag": "volts",
                "CoefA": "1.000000",
                "CoefB": "0.000000",
                "Unit": "V",
                "Action": "4",
            }
        ]
        # Calling write_output_csv with positional string parameters: output, rows, manufacturer, model, protocol, category, forced_write
        Generator.write_output_csv(
            buf, rows, "SampleMfg", "SampleModel", "modbusRTU", "Inverter", ""
        )
        res = buf.getvalue()
        self.assertIn("modbusRTU;Inverter;SampleMfg;SampleModel", res)
        self.assertIn("1;3;40001;U16;;Volts;volts;1.000000;0.000000;V;4", res)

    def test_run_generator_non_existent_input_file(self):
        cfg = GeneratorConfig(input_file="non_existent_file_abc_123.csv", output="out.csv")
        run_generator(cfg)  # Should gracefully log error without crashing

    def test_extractor_csv_corrupt_text(self):
        extractor = Extractor()
        with tempfile.NamedTemporaryFile("wb", suffix=".csv", delete=False) as f:
            f.write(b"\xff\xfe\x00\x00corrupt_binary_data_without_delimiters")
            f_path = f.name
        try:
            tables = list(extractor.extract_from_csv(f_path))
            self.assertEqual(len(tables), 1)
            rows = list(tables[0])
            self.assertEqual(len(rows), 0)
        finally:
            if os.path.exists(f_path):
                os.remove(f_path)

    def test_extractor_xml_missing_file(self):
        extractor = Extractor()
        tables = list(extractor.extract_from_xml("non_existent_xml_file.xml"))
        self.assertEqual(len(tables), 1)
        rows = list(tables[0])
        self.assertEqual(len(rows), 0)


if __name__ == "__main__":
    unittest.main()
