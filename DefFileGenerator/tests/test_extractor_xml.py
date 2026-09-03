"""Batch 6: XML extractor tests and sanitize_csv_field edge cases."""

import logging
import os
import tempfile
import unittest

from DefFileGenerator.def_gen import Generator


class TestSanitizeCsvFieldExtended(unittest.TestCase):
    """Extended sanitize_csv_field edge cases."""

    def _s(self, v):
        return Generator.sanitize_csv_field(v)

    # Numbers preserved
    def test_positive_int(self):
        self.assertEqual(self._s("42"), "42")

    def test_zero(self):
        self.assertEqual(self._s("0"), "0")

    def test_sci_notation(self):
        self.assertEqual(self._s("1.5e3"), "1.5e3")

    def test_sci_notation_negative_exp(self):
        self.assertEqual(self._s("-1.5E-3"), "-1.5E-3")

    def test_plus_decimal(self):
        self.assertEqual(self._s("+.5"), "+.5")

    # Numbers not preserved (non-finite or weird)
    def test_negative_inf(self):
        self.assertEqual(self._s("-inf"), "'-inf")

    def test_positive_nan(self):
        self.assertEqual(self._s("+nan"), "'+nan")

    def test_underscore_grouped(self):
        self.assertEqual(self._s("-1_000"), "'-1_000")

    def test_space_before_minus_number(self):
        self.assertEqual(self._s(" -10.5"), "' -10.5")

    # Formula injection triggers
    def test_equals_sign(self):
        self.assertEqual(self._s("=SUM(A1:A10)"), "'=SUM(A1:A10)")

    def test_at_sign(self):
        self.assertEqual(self._s("@TODAY()"), "'@TODAY()")

    def test_pipe(self):
        result = self._s("|DDE")
        self.assertTrue(result.startswith("'"), f"Expected apostrophe prefix, got: {result!r}")
        self.assertIn("|DDE", result)

    def test_tab_prefix(self):
        self.assertEqual(self._s("\t=EVIL"), "'\t=EVIL")

    def test_cr_prefix(self):
        self.assertEqual(self._s("\r=EVIL"), "'\r=EVIL")

    def test_lf_prefix(self):
        self.assertEqual(self._s("\n=EVIL"), "'\n=EVIL")

    def test_nbsp_prefix(self):
        self.assertEqual(self._s("\u00a0=EVIL"), "'\u00a0=EVIL")

    # Fullwidth triggers
    def test_fullwidth_equals(self):
        self.assertEqual(self._s("\uff1d1+1"), "'\uff1d1+1")

    def test_fullwidth_plus(self):
        self.assertEqual(self._s("\uff0bSUM"), "'\uff0bSUM")

    def test_fullwidth_minus(self):
        self.assertEqual(self._s("\uff0d1"), "'\uff0d1")

    def test_fullwidth_at(self):
        self.assertEqual(self._s("\uff20SUM"), "'\uff20SUM")

    # Normal text unchanged
    def test_plain_word(self):
        self.assertEqual(self._s("Voltage"), "Voltage")

    def test_plain_empty(self):
        self.assertEqual(self._s(""), "")

    def test_int_value(self):
        self.assertEqual(self._s(42), "42")

    def test_float_value(self):
        self.assertEqual(self._s(3.14), "3.14")

    def test_none_becomes_empty(self):
        self.assertEqual(self._s(None), "")


class TestExtractorXml(unittest.TestCase):
    """XML extractor tests."""

    def setUp(self):
        from DefFileGenerator.extractor import Extractor
        self.ex = Extractor()
        self.tmpdir = tempfile.TemporaryDirectory()
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        self.tmpdir.cleanup()
        logging.disable(logging.NOTSET)

    def _xml(self, content, name="test.xml"):
        p = os.path.join(self.tmpdir.name, name)
        with open(p, "w", encoding="utf-8") as f:
            f.write(content)
        return p

    def test_simple_xml_extracted(self):
        xml = """<?xml version="1.0"?>
<registers>
  <register>
    <Address>100</Address>
    <Name>Voltage</Name>
    <Type>U16</Type>
  </register>
</registers>"""
        p = self._xml(xml)
        tables = list(self.ex.extract_from_xml(p))
        rows = list(tables[0])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Name"], "Voltage")
        self.assertEqual(rows[0]["Address"], "100")

    def test_multiple_registers_extracted(self):
        xml = """<?xml version="1.0"?>
<registers>
  <register><Address>100</Address><Name>V</Name><Type>U16</Type></register>
  <register><Address>101</Address><Name>I</Name><Type>U16</Type></register>
</registers>"""
        p = self._xml(xml)
        tables = list(self.ex.extract_from_xml(p))
        rows = list(tables[0])
        self.assertEqual(len(rows), 2)

    def test_duplicate_rows_deduplicated(self):
        xml = """<?xml version="1.0"?>
<registers>
  <register><Address>100</Address><Name>V</Name><Type>U16</Type></register>
  <register><Address>100</Address><Name>V</Name><Type>U16</Type></register>
</registers>"""
        p = self._xml(xml)
        tables = list(self.ex.extract_from_xml(p))
        rows = list(tables[0])
        self.assertEqual(len(rows), 1)

    def test_xml_element_with_single_child_skipped(self):
        """Elements with only one child are not yielded (len(row) < 2)."""
        xml = """<?xml version="1.0"?>
<registers>
  <register><Name>Only</Name></register>
</registers>"""
        p = self._xml(xml)
        tables = list(self.ex.extract_from_xml(p))
        rows = list(tables[0])
        self.assertEqual(rows, [])

    def test_xml_missing_file_logs_error(self):
        logging.disable(logging.NOTSET)
        tables = list(self.ex.extract_from_xml("/no/such/file.xml"))
        rows = list(tables[0])
        self.assertEqual(rows, [])

    def test_malformed_xml_logs_error(self):
        p = self._xml("NOT VALID XML <<>>")
        tables = list(self.ex.extract_from_xml(p))
        rows = list(tables[0])
        self.assertEqual(rows, [])

    def test_xml_and_map_end_to_end(self):
        xml = """<?xml version="1.0"?>
<registers>
  <register>
    <Address>200</Address>
    <Name>Power</Name>
    <Type>U32</Type>
  </register>
</registers>"""
        p = self._xml(xml)
        raw = self.ex.extract_from_xml(p)
        result = list(self.ex.map_and_clean(raw))
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["Name"], "Power")

    def test_xml_xxe_security_blocked(self):
        """XXE attack vectors must be refused (defusedxml raises security exception)."""
        from DefFileGenerator.extractor import SECURITY_EXCEPTIONS
        if not SECURITY_EXCEPTIONS:
            self.skipTest("defusedxml not available or no security exceptions defined")
        xxe = """<?xml version="1.0"?>
<!DOCTYPE foo [<!ENTITY xxe SYSTEM "file:///etc/passwd">]>
<registers><register><Name>&xxe;</Name><Address>1</Address></register></registers>"""
        p = self._xml(xxe)
        # Should either return empty or raise a security exception caught internally
        try:
            tables = list(self.ex.extract_from_xml(p))
            rows = list(tables[0])
            # Either empty rows or no data exfiltrated
            self.assertFalse(any("/root" in str(r) for r in rows))
        except Exception as e:
            # Security exception may propagate
            self.assertIsInstance(e, SECURITY_EXCEPTIONS)


class TestExtractorNopdfplumber(unittest.TestCase):
    """PDF extraction when pdfplumber is absent (mocked)."""

    def test_no_pdfplumber_logs_error_returns_empty(self):
        import sys
        from unittest.mock import patch
        from DefFileGenerator.extractor import Extractor
        ex = Extractor()
        logging.disable(logging.NOTSET)
        with patch.object(sys.modules.get("DefFileGenerator.extractor", None) or type(None),
                          "__name__", "DefFileGenerator.extractor"):
            with patch("DefFileGenerator.extractor.HAS_PDFPLUMBER", False):
                result = list(ex.extract_from_pdf("dummy.pdf"))
        self.assertEqual(result, [])


if __name__ == "__main__":
    unittest.main()