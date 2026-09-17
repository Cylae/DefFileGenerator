import io
import tempfile
import unittest

from fastapi.testclient import TestClient

import DefFileGenerator
from DefFileGenerator import Generator, WebdynDefConfig
from generate_webdyn_def import WebdynDefConfig as RootWebdynDefConfig
from web.app import MAX_WEB_REGISTERS, app


class TestWave9Hardening(unittest.TestCase):
    def setUp(self):
        self.client = TestClient(app)
        self.tmpdir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.tmpdir.cleanup()

    def test_webdyn_def_config_contract_and_exports(self):
        """Verify WebdynDefConfig is exported and matches root script definition."""
        self.assertIn("WebdynDefConfig", DefFileGenerator.__all__)
        self.assertIs(WebdynDefConfig, RootWebdynDefConfig)

        cfg = WebdynDefConfig(
            input_file="input.xlsx",
            output_file="output.csv",
            manufacturer="VendorA",
            model="ModelB",
        )
        self.assertEqual(cfg.input_file, "input.xlsx")
        self.assertEqual(cfg.output_file, "output.csv")
        self.assertEqual(cfg.manufacturer, "VendorA")
        self.assertEqual(cfg.model, "ModelB")
        self.assertEqual(cfg.protocol, "modbusRTU")
        self.assertEqual(cfg.category, "Inverter")
        self.assertEqual(cfg.address_offset, 0)
        self.assertTrue(cfg.strict_validation)

    def test_csv_injection_unicode_whitespace_and_zero_width_bypasses(self):
        """Verify sanitize_csv_field escapes formula triggers prefixed by all Unicode space and invisible chars."""
        unicode_prefixes = [
            "\u2000",  # En Quad
            "\u2001",  # Em Quad
            "\u2002",  # En Space
            "\u2003",  # Em Space
            "\u2004",  # Three-Per-Em Space
            "\u2005",  # Four-Per-Em Space
            "\u2006",  # Six-Per-Em Space
            "\u2007",  # Figure Space
            "\u2008",  # Punctuation Space
            "\u2009",  # Thin Space
            "\u200a",  # Hair Space
            "\u200b",  # Zero Width Space
            "\u200c",  # Zero Width Non-Joiner
            "\u200d",  # Zero Width Joiner
            "\u200e",  # Left-to-Right Mark
            "\u200f",  # Right-to-Left Mark
            "\u2028",  # Line Separator
            "\u2029",  # Paragraph Separator
            "\u202f",  # Narrow No-Break Space
            "\u205f",  # Medium Mathematical Space
            "\u2060",  # Word Joiner
            "\u3000",  # Ideographic Space
            "\u00a0",  # Non-Breaking Space
            "\u0085",  # Next Line
            "\u00ad",  # Soft Hyphen
            "\u180e",  # Mongolian Vowel Separator
            "\ufeff",  # Zero Width No-Break Space / BOM
        ]

        formula_triggers = ["=1+1", "+1+1", "-1+1", "@SUM(1,1)", "|cmd", "%COMSPEC%"]

        for prefix in unicode_prefixes:
            for trigger in formula_triggers:
                payload = f"{prefix}{trigger}"
                sanitized = Generator.sanitize_csv_field(payload)
                self.assertTrue(
                    sanitized.startswith("'"),
                    f"Prefix {repr(prefix)} with trigger {trigger} was not escaped! Result: {repr(sanitized)}",
                )

    def test_embedded_industrial_types_normalization(self):
        """Verify C/embedded Modbus types in vendor documentation are recognized and normalized."""
        test_cases = [
            ("WORD", "U16"),
            ("word", "U16"),
            ("uword", "U16"),
            ("DWORD", "U32"),
            ("dword", "U32"),
            ("udword", "U32"),
            ("dword swap", "U32_WB"),
            ("QWORD", "U64"),
            ("qword", "U64"),
            ("BYTE", "U8"),
            ("byte", "U8"),
            ("ubyte", "U8"),
            ("unsigned long", "U32"),
            ("unsigned long int", "U32"),
            ("ulong", "U32"),
            ("signed long", "I32"),
            ("signed long int", "I32"),
            ("long", "I32"),
            ("slong", "I32"),
            ("unsigned short", "U16"),
            ("ushort", "U16"),
            ("signed short", "I16"),
            ("short", "I16"),
            ("sshort", "I16"),
            ("unsigned char", "U8"),
            ("uchar", "U8"),
            ("signed char", "I8"),
            ("schar", "I8"),
            ("uint32_t", "U32"),
            ("int32_t", "I32"),
            ("uint16_t", "U16"),
            ("uint8_t", "U8"),
        ]
        for raw, expected in test_cases:
            normalized = Generator.normalize_type(raw)
            self.assertEqual(
                normalized,
                expected,
                f"normalize_type('{raw}') returned '{normalized}', expected '{expected}'",
            )
            self.assertTrue(
                Generator.validate_type(normalized),
                f"Normalized type '{normalized}' failed validate_type!",
            )

    def test_trailing_dot_address_normalization(self):
        """Verify trailing period in address cells (from OCR or table numbering) is safely stripped."""
        self.assertEqual(Generator.normalize_address_val("1000."), "1000")
        self.assertEqual(Generator.normalize_address_val("1001."), "1001")
        self.assertEqual(Generator.normalize_address_val("0x1000."), "4096")
        # Ensure IP addresses or normal dotted values are unaffected
        self.assertEqual(Generator.normalize_address_val("1.2.3.4"), "1.2.3.4")

    def test_web_validate_file_extension_rejection(self):
        """Verify /api/validate rejects non-CSV/non-text file uploads."""
        file_obj = io.BytesIO(b"binary payload")
        response = self.client.post(
            "/api/validate",
            files={"file": ("malicious.exe", file_obj, "application/octet-stream")},
        )
        self.assertEqual(response.status_code, 400)
        self.assertIn("Validation requires a CSV definition file", response.json()["detail"])

        file_pdf = io.BytesIO(b"%PDF-1.4...")
        response_pdf = self.client.post(
            "/api/validate",
            files={"file": ("documentation.pdf", file_pdf, "application/pdf")},
        )
        self.assertEqual(response_pdf.status_code, 400)
        self.assertIn("Validation requires a CSV definition file", response_pdf.json()["detail"])

    def test_web_convert_max_register_limit(self):
        """Verify /api/convert rejects uploads exceeding MAX_WEB_REGISTERS."""
        # Generate a CSV that contains more than MAX_WEB_REGISTERS registers
        lines = ["Register,Name,Data Type,Unit,Scale,Access"]
        for i in range(MAX_WEB_REGISTERS + 5):
            lines.append(f"{i},Var_{i},uint16,V,1,R")
        big_csv = "\n".join(lines).encode("utf-8")

        response = self.client.post(
            "/api/convert",
            files={"file": ("huge.csv", io.BytesIO(big_csv), "text/csv")},
            data={"manufacturer": "BigMfg", "model": "BigModel"},
        )
        self.assertEqual(response.status_code, 400)
        self.assertIn("exceeding the maximum allowable limit", response.json()["detail"])


if __name__ == "__main__":
    unittest.main()
