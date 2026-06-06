import unittest
import os
import csv
import json

try:
    from openpyxl import Workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from reportlab.pdfgen import canvas
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
    from reportlab.lib.pagesizes import letter
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

import unittest.mock
from DefFileGenerator.extractor import Extractor

class TestExtractor(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()
        self.excel_file = "test_registers.xlsx"
        self.pdf_file = "test_registers.pdf"
        self.mapping_file = "test_mapping.json"

        # Create dummy Excel
        if HAS_OPENPYXL:
            wb = Workbook()
            ws = wb.active
            ws.title = "Registers"
            ws.append(["Reg Addr", "Description", "Data Type", "Unit"])
            ws.append(["0x0001", "Voltage", "Uint16", "V"])
            ws.append(["0x0002", "Current", "Int32", "A"])
            ws.append(["40001", "Power", "Float32", "W"])
            wb.save(self.excel_file)

        # Create dummy PDF
        if HAS_REPORTLAB:
            doc = SimpleDocTemplate(self.pdf_file, pagesize=letter)
            data = [
                ["Address", "Name", "Type"],
                ["1000", "Temp", "U16"],
                ["1001", "Humid", "U16"]
            ]
            t = Table(data)
            t.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, (0, 0, 0)),
            ]))
            elements = [t]
            doc.build(elements)

    def tearDown(self):
        for f in [self.excel_file, self.pdf_file, self.mapping_file]:
            if os.path.exists(f):
                os.remove(f)

    def test_normalize_type(self):
        # normalize_type is now a static method
        self.assertEqual(Extractor.normalize_type("Uint16"), "U16")
        self.assertEqual(Extractor.normalize_type("Int32"), "I32")
        self.assertEqual(Extractor.normalize_type("Float32"), "F32")
        self.assertEqual(Extractor.normalize_type("unsigned int 16"), "U16")

    @unittest.skipUnless(HAS_OPENPYXL, "openpyxl not installed")
    def test_extract_from_excel(self):
        if not os.path.exists(self.excel_file): self.skipTest("Excel file not created")
        data = [list(table) for table in self.extractor.extract_from_excel(self.excel_file)]
        self.assertEqual(len(data), 1) # One sheet = one table
        self.assertEqual(len(data[0]), 3) # 3 data rows
        self.assertEqual(str(data[0][0]["Reg Addr"]), "0x0001")

    @unittest.skipUnless(HAS_OPENPYXL, "openpyxl not installed")
    def test_map_and_clean_excel(self):
        if not os.path.exists(self.excel_file): self.skipTest("Excel file not created")
        raw_data = self.extractor.extract_from_excel(self.excel_file)
        # Custom mapping
        self.extractor.mapping = {
            "Address": "Reg Addr",
            "Name": "Description",
            "Type": "Data Type"
        }
        mapped = list(self.extractor.map_and_clean(raw_data))
        self.assertEqual(len(mapped), 3)
        self.assertEqual(mapped[0]["Address"], "1")
        self.assertEqual(mapped[0]["Name"], "Voltage")
        self.assertEqual(mapped[0]["Type"], "U16")
        self.assertEqual(mapped[1]["Type"], "I32")
        self.assertEqual(mapped[2]["Type"], "F32")

    @unittest.skipUnless(HAS_REPORTLAB, "reportlab not installed")
    def test_extract_from_pdf(self):
        if not os.path.exists(self.pdf_file): self.skipTest("PDF file not created")
        data = [list(table) for table in self.extractor.extract_from_pdf(self.pdf_file)]
        self.assertEqual(len(data), 1) # One table found
        self.assertEqual(len(data[0]), 2) # 2 data rows
        self.assertEqual(data[0][0]["Address"], "1000")
        self.assertEqual(data[0][0]["Name"], "Temp")

    def test_fuzzy_mapping(self):
        # Even without explicit mapping, it should find Name, Address, Type if headers are similar
        raw_data = [
            [{"Register Address": "0x10", "Variable Name": "Test", "Data Type": "Uint16"}]
        ]
        mapped = list(self.extractor.map_and_clean(raw_data))
        self.assertEqual(mapped[0]["Address"], "16")
        self.assertEqual(mapped[0]["Name"], "Test")
        self.assertEqual(mapped[0]["Type"], "U16")

    def test_pdf_page_validation(self):
        # Mock pdfplumber to test page range validation
        mock_pdfplumber = unittest.mock.MagicMock()
        with unittest.mock.patch.dict('sys.modules', {'pdfplumber': mock_pdfplumber}):
            mock_pdf = unittest.mock.MagicMock()
            mock_pdf.pages = [unittest.mock.MagicMock()] * 10
            mock_pdfplumber.open.return_value.__enter__.return_value = mock_pdf

            # Use patch to set HAS_PDFPLUMBER and pdfplumber in the module
            with unittest.mock.patch('DefFileGenerator.extractor.HAS_PDFPLUMBER', True), \
                 unittest.mock.patch('DefFileGenerator.extractor.pdfplumber', mock_pdfplumber, create=True):
                extractor = Extractor()
                # Valid pages
                list(extractor.extract_from_pdf("dummy.pdf", pages="1,3,5"))

                # Out of range pages
                with self.assertLogs(level='WARNING') as cm:
                    list(extractor.extract_from_pdf("dummy.pdf", pages="1,15"))
                    self.assertTrue(any("Page 15 out of range" in msg for msg in cm.output))

if __name__ == "__main__":
    unittest.main()
