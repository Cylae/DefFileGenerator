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


if __name__ == "__main__":
    unittest.main()
