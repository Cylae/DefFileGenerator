import unittest
import logging
import io
import os
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator
from DefFileGenerator.extractor import Extractor

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.extractor = Extractor()
        # Suppress logging during tests
        logging.getLogger().setLevel(logging.CRITICAL)

    def test_address_range_validation(self):
        # Valid address
        self.assertTrue(self.generator.validate_address("100", "U16"))

        # Out of range address (should still return True for format, but log warning)
        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_address("70000", "U16"))
            self.assertIn("Address 70000 is outside standard Modbus range", cm.output[0])

        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_address("-1", "U16"))
            self.assertIn("Address -1 is outside standard Modbus range", cm.output[0])

    def test_intelligent_action_defaulting(self):
        rows = [
            {'Name': 'Holding', 'Address': '1', 'Type': 'U16', 'RegisterType': 'Holding Register'},
            {'Name': 'Input', 'Address': '2', 'Type': 'U16', 'RegisterType': 'Input Register'},
            {'Name': 'Coil', 'Address': '3', 'Type': 'U16', 'RegisterType': 'Coil'},
            {'Name': 'Discrete', 'Address': '4', 'Type': 'U16', 'RegisterType': 'Discrete Input'},
        ]
        # RegisterType mapping: Coil=1, Discrete=2, Holding=3, Input=4
        processed = list(self.generator.process_rows(rows))

        self.assertEqual(len(processed), 4)
        # Holding Register -> Action 1
        self.assertEqual(processed[0]['Action'], '1')
        # Input Register -> Action 4
        self.assertEqual(processed[1]['Action'], '4')
        # Coil -> Action 1
        self.assertEqual(processed[2]['Action'], '1')
        # Discrete Input -> Action 4
        self.assertEqual(processed[3]['Action'], '4')

    def test_pdf_page_range_validation(self):
        # This test requires pdfplumber or a mock
        try:
            import pdfplumber
            from reportlab.pdfgen import canvas

            pdf_path = "test_pages.pdf"
            c = canvas.Canvas(pdf_path)
            c.drawString(100, 750, "Page 1")
            c.showPage()
            c.save()

            # Page 1 (valid)
            raw = list(self.extractor.extract_from_pdf(pdf_path, pages=[1]))
            self.assertEqual(len(raw), 0) # No tables, but shouldn't error

            # Page 2 (out of range)
            with self.assertLogs(level='WARNING') as cm:
                raw = list(self.extractor.extract_from_pdf(pdf_path, pages=[2]))
                self.assertIn("Page 2 out of range", cm.output[0])

            if os.path.exists(pdf_path): os.remove(pdf_path)
        except ImportError:
            self.skipTest("pdfplumber or reportlab not installed")

    def test_excel_sheet_handling(self):
        try:
            import openpyxl
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "RealSheet"
            wb.save("test_sheets.xlsx")

            # Valid sheet
            raw = list(self.extractor.extract_from_excel("test_sheets.xlsx", sheet_name="RealSheet"))
            self.assertEqual(len(raw), 1)

            # Missing sheet
            with self.assertLogs(level='WARNING') as cm:
                raw = list(self.extractor.extract_from_excel("test_sheets.xlsx", sheet_name="FakeSheet"))
                self.assertIn("Sheet 'FakeSheet' not found", cm.output[0])
                self.assertEqual(len(raw), 0)

            if os.path.exists("test_sheets.xlsx"): os.remove("test_sheets.xlsx")
        except ImportError:
            self.skipTest("openpyxl not installed")

    def test_stats_logging(self):
        rows = [
            {'Info1': '3', 'Info2': '1', 'Info3': 'U16', 'Info4': '', 'Name': 'N1', 'Tag': 'T1', 'CoefA': '1.0', 'CoefB': '0.0', 'Unit': 'V', 'Action': '4'},
            {'Info1': '4', 'Info2': '2', 'Info3': 'U16', 'Info4': '', 'Name': 'N2', 'Tag': 'T2', 'CoefA': '1.0', 'CoefB': '0.0', 'Unit': 'A', 'Action': '4'},
        ]
        out = io.StringIO()
        with self.assertLogs(level='INFO') as cm:
            self.generator.write_output_csv(out, rows, "MFG", "MOD")
            self.assertTrue(any("Processed registers - Holding Registers: 1, Input Registers: 1" in line for line in cm.output))

if __name__ == '__main__':
    unittest.main()
