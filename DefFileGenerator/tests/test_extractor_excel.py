"""Batch 7: Excel extractor and extractor module CLI tests."""

import logging
import os
import sys
import tempfile
import unittest


class TestExtractorExcel(unittest.TestCase):
    """Tests for extract_from_excel using real xlsx files via openpyxl."""

    def setUp(self):
        from DefFileGenerator.extractor import Extractor, HAS_OPENPYXL
        if not HAS_OPENPYXL:
            self.skipTest("openpyxl not available")
        self.ex = Extractor()
        self.tmpdir = tempfile.TemporaryDirectory()
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        self.tmpdir.cleanup()
        logging.disable(logging.NOTSET)

    def _make_excel(self, sheets, filename="test.xlsx"):
        import openpyxl
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # remove default sheet
        for sheet_name, rows in sheets.items():
            ws = wb.create_sheet(sheet_name)
            for row in rows:
                ws.append(row)
        p = os.path.join(self.tmpdir.name, filename)
        wb.save(p)
        return p

    def test_simple_excel_single_sheet(self):
        p = self._make_excel({"Sheet1": [
            ["Name", "Address", "Type"],
            ["Voltage", "100", "U16"],
        ]})
        tables = list(self.ex.extract_from_excel(p))
        rows = list(tables[0])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Name"], "Voltage")

    def test_excel_multiple_sheets_yields_multiple_tables(self):
        p = self._make_excel({
            "Sheet1": [["Name", "Address", "Type"], ["V1", "100", "U16"]],
            "Sheet2": [["Name", "Address", "Type"], ["V2", "200", "U32"]],
        })
        tables = list(self.ex.extract_from_excel(p))
        self.assertEqual(len(tables), 2)
        rows1 = list(tables[0])
        rows2 = list(tables[1])
        self.assertEqual(rows1[0]["Name"], "V1")
        self.assertEqual(rows2[0]["Name"], "V2")

    def test_excel_specific_sheet_by_name(self):
        p = self._make_excel({
            "Sheet1": [["Name", "Address", "Type"], ["V1", "100", "U16"]],
            "RegisterData": [["Name", "Address", "Type"], ["V2", "200", "U16"]],
        })
        tables = list(self.ex.extract_from_excel(p, sheet_name="RegisterData"))
        rows = list(tables[0])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Name"], "V2")

    def test_excel_invalid_sheet_name_logs_error(self):
        logging.disable(logging.NOTSET)
        p = self._make_excel({"Sheet1": [["Name", "Address"], ["V1", "100"]]})
        tables = list(self.ex.extract_from_excel(p, sheet_name="NoSuchSheet"))
        # Generator created, but iterating should log error
        with self.assertLogs(level="ERROR"):
            rows = list(tables[0])
        self.assertEqual(rows, [])

    def test_excel_empty_sheet(self):
        p = self._make_excel({"Empty": []})
        tables = list(self.ex.extract_from_excel(p))
        rows = list(tables[0])
        self.assertEqual(rows, [])

    def test_excel_header_only_sheet(self):
        p = self._make_excel({"Data": [["Name", "Address", "Type"]]})
        tables = list(self.ex.extract_from_excel(p))
        rows = list(tables[0])
        self.assertEqual(rows, [])

    def test_excel_missing_file_logs_error(self):
        logging.disable(logging.NOTSET)
        tables = list(self.ex.extract_from_excel("/no/such/file.xlsx"))
        with self.assertLogs(level="ERROR"):
            rows = list(tables[0])
        self.assertEqual(rows, [])

    def test_excel_skips_fully_blank_rows(self):
        p = self._make_excel({"Data": [
            ["Name", "Address", "Type"],
            ["V1", "100", "U16"],
            [None, None, None],   # blank row
            ["V2", "200", "U16"],
        ]})
        tables = list(self.ex.extract_from_excel(p))
        rows = list(tables[0])
        self.assertEqual(len(rows), 2)

    def test_excel_and_map_end_to_end(self):
        p = self._make_excel({"Data": [
            ["Name", "Address", "Type"],
            ["GridV", "100", "U16"],
        ]})
        raw = self.ex.extract_from_excel(p)
        result = list(self.ex.map_and_clean(raw))
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["Name"], "GridV")

    def test_excel_not_openpyxl_logs_error(self):
        from unittest.mock import patch
        logging.disable(logging.NOTSET)
        with patch("DefFileGenerator.extractor.HAS_OPENPYXL", False):
            tables = list(self.ex.extract_from_excel("dummy.xlsx"))
        self.assertEqual(tables, [])


class TestExtractorModuleCli(unittest.TestCase):
    """Tests for extractor.main() CLI entry point (lines 430–484)."""

    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.tmpdir.cleanup()

    def _csv(self, content="Name,Address,Type\nVar1,100,U16\n"):
        p = os.path.join(self.tmpdir.name, "in.csv")
        with open(p, "w", encoding="utf-8") as f:
            f.write(content)
        return p

    def _xml(self, content):
        p = os.path.join(self.tmpdir.name, "in.xml")
        with open(p, "w", encoding="utf-8") as f:
            f.write(content)
        return p

    def test_extractor_cli_csv_to_stdout(self):
        import io
        from unittest.mock import patch
        src = self._csv()
        with patch("sys.argv", ["extractor", src]):
            captured = io.StringIO()
            with patch("sys.stdout", captured):
                from DefFileGenerator.extractor import main
                main()
        self.assertIn("Name", captured.getvalue())

    def test_extractor_cli_csv_to_file(self):
        src = self._csv()
        out = os.path.join(self.tmpdir.name, "out.csv")
        from unittest.mock import patch
        with patch("sys.argv", ["extractor", src, "-o", out]):
            from DefFileGenerator.extractor import main
            main()
        self.assertTrue(os.path.exists(out))

    def test_extractor_cli_unsupported_ext_exits_1(self):
        p = os.path.join(self.tmpdir.name, "file.docx")
        with open(p, "w") as f:
            f.write("dummy")
        from unittest.mock import patch
        with patch("sys.argv", ["extractor", p]):
            with self.assertRaises(SystemExit) as cm:
                from DefFileGenerator.extractor import main
                main()
        self.assertEqual(cm.exception.code, 1)

    def test_extractor_cli_xml_to_stdout(self):
        xml_content = """<?xml version="1.0"?>
<registers>
  <register><Address>100</Address><Name>Power</Name><Type>U32</Type></register>
</registers>"""
        src = self._xml(xml_content)
        import io
        from unittest.mock import patch
        with patch("sys.argv", ["extractor", src]):
            captured = io.StringIO()
            with patch("sys.stdout", captured):
                from DefFileGenerator.extractor import main
                main()
        self.assertIn("Name", captured.getvalue())

    def test_extractor_cli_with_address_offset(self):
        src = self._csv("Name,Address,Type\nVar1,100,U16\n")
        out = os.path.join(self.tmpdir.name, "out.csv")
        from unittest.mock import patch
        with patch("sys.argv", ["extractor", src, "-o", out, "--address-offset", "10"]):
            from DefFileGenerator.extractor import main
            main()
        with open(out) as f:
            content = f.read()
        self.assertIn("110", content)


if __name__ == "__main__":
    unittest.main()