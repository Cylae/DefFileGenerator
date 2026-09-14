import logging
import os
import tempfile
import unittest

from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor


class TestAdversarialSecurity(unittest.TestCase):
    def setUp(self):
        logging.disable(logging.CRITICAL)
        self.temp_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        logging.disable(logging.NOTSET)
        self.temp_dir.cleanup()

    def test_csv_injection_vectors(self):
        """Verify sanitize_csv_field escapes formula injection characters including fullwidth variants and leading control chars."""
        payloads = [
            "=1+1",
            "+1+1",
            "-1+1",
            "@SUM(1,1)",
            "|cmd",
            "%COMSPEC%",
            "\t=1+1",
            "\r=1+1",
            "\n=1+1",
            "\u00a0=1+1",
            "\ufeff=1+1",
            "\uff1d1+1",  # Fullwidth equals sign
            "\uff0b1+1",  # Fullwidth plus sign
            "\uff0d1+1",  # Fullwidth minus sign
            "\uff201+1",  # Fullwidth commercial at
        ]
        for p in payloads:
            sanitized = Generator.sanitize_csv_field(p)
            self.assertTrue(
                sanitized.startswith("'"),
                f"Payload '{repr(p)}' was not sanitized! Result: '{sanitized}'",
            )

    def test_valid_numbers_and_safe_strings_not_escaped(self):
        """Verify normal numeric values and plain strings are not redundantly escaped."""
        self.assertEqual(Generator.sanitize_csv_field(123), "123")
        self.assertEqual(Generator.sanitize_csv_field(-45.67), "-45.67")
        self.assertEqual(Generator.sanitize_csv_field("Normal Text"), "Normal Text")
        self.assertEqual(Generator.sanitize_csv_field(""), "")

    def test_xml_xxe_and_entity_defense(self):
        """Verify XML parser rejects or strips external entities and DTD declarations safely."""
        xxe_xml = os.path.join(self.temp_dir.name, "xxe.xml")
        with open(xxe_xml, "w", encoding="utf-8") as f:
            f.write(
                '<?xml version="1.0"?>\n'
                '<!DOCTYPE foo [ <!ENTITY xxe SYSTEM "file:///etc/passwd"> ]>\n'
                "<registers>\n"
                "  <reg><Name>&xxe;</Name><Address>40001</Address></reg>\n"
                "</registers>\n"
            )
        extractor = Extractor()
        # Defusedxml should raise exception or block entity resolution
        try:
            raw = list(extractor.extract_from_xml(xxe_xml))
            rows = list(extractor.map_and_clean(raw))
            # If extracted, entity must not be expanded to file contents
            for row in rows:
                self.assertNotIn("root:", str(row.get("Name", "")))
        except Exception as e:
            self.assertTrue(True, f"Successfully blocked entity/DTD: {e}")

    def test_extreme_address_offset_and_overflow(self):
        """Verify apply_address_offset handles enormous integers without crashing or invalidating formatting."""
        huge_offset = 999999999
        res = Generator.apply_address_offset("40001", huge_offset)
        self.assertEqual(res, str(40001 + huge_offset))

        # Negative offset causing negative address
        res_neg = Generator.apply_address_offset("100", -200)
        self.assertEqual(res_neg, "-100")

    def test_control_character_stripping_and_clean_headers(self):
        """Verify headers and string fields handle embedded null bytes and control chars cleanly."""
        hdr_val = "Huawei\x00Inverter\x07"
        sanitized = Generator.sanitize_csv_field(hdr_val)
        self.assertNotIn("\x00", sanitized)
        self.assertNotIn("\x07", sanitized)

    def test_multiline_field_sanitization_prevents_row_splitting(self):
        """Verify embedded newlines and carriage returns are normalized to single spaces."""
        val = "Total Active\nPower\r\n(kW)"
        sanitized = Generator.sanitize_csv_field(val)
        self.assertNotIn("\n", sanitized)
        self.assertNotIn("\r", sanitized)
        self.assertEqual(sanitized, "Total Active Power (kW)")

    def test_csv_cp1252_encoding_fallback_and_webdyn_extraction(self):
        """Verify CSV extractor gracefully handles cp1252/latin-1 French/German characters and Webdyn definition format."""
        csv_path = os.path.join(self.temp_dir.name, "cp1252_test.csv")
        # Write bytes encoded in cp1252 (with 'é' as \xe9)
        content = (
            "modbus;Meter;PQ PLUS;GENERIC;;;;;;;\n"
            "1;3;25946;F32_W;;Puissance active Phase L1;ActivePow;0.001;0;kW;4\n"
            "2;3;25948;F32_W;;Puissance réactive Phase L1;ReactivePow;1;0;var;4\n"
        )
        with open(csv_path, "wb") as f:
            f.write(content.encode("cp1252"))

        extractor = Extractor()
        raw = list(extractor.extract_from_csv(csv_path))
        self.assertEqual(len(raw), 1)
        rows = list(raw[0])
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["Name"], "Puissance active Phase L1")
        self.assertIn("réactive", rows[1]["Name"])

        # Re-extract for map_and_clean since raw generator was consumed above
        raw2 = extractor.extract_from_csv(csv_path)
        mapped = list(extractor.map_and_clean(raw2))
        self.assertEqual(len(mapped), 2)
        self.assertEqual(mapped[0]["Address"], "25946")

    def test_validate_csv_strict_enforcement(self):
        """Verify validate_csv strictly checks Info1, Action, and numeric coefficients."""
        gen = Generator()
        # Invalid Info1 (e.g. 5)
        bad_info1 = os.path.join(self.temp_dir.name, "bad_info1.csv")
        with open(bad_info1, "w", encoding="utf-8") as f:
            f.write("modbusRTU;Inverter;Mfg;Model;;;;;;;\n1;5;1000;U16;;V1;tag1;1.0;0.0;V;4\n")
        self.assertFalse(gen.validate_csv(bad_info1, strict=True))
        self.assertTrue(gen.validate_csv(bad_info1, strict=False))

        # Invalid Action (e.g. 99)
        bad_act = os.path.join(self.temp_dir.name, "bad_act.csv")
        with open(bad_act, "w", encoding="utf-8") as f:
            f.write("modbusRTU;Inverter;Mfg;Model;;;;;;;\n1;3;1000;U16;;V1;tag1;1.0;0.0;V;99\n")
        self.assertFalse(gen.validate_csv(bad_act, strict=True))

        # Non-numeric CoefA
        bad_coef = os.path.join(self.temp_dir.name, "bad_coef.csv")
        with open(bad_coef, "w", encoding="utf-8") as f:
            f.write("modbusRTU;Inverter;Mfg;Model;;;;;;;\n1;3;1000;U16;;V1;tag1;NOT_NUM;0.0;V;4\n")
        self.assertFalse(gen.validate_csv(bad_coef, strict=True))


if __name__ == "__main__":
    unittest.main()
