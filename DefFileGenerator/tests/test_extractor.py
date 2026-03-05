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

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

from DefFileGenerator.extractor import Extractor

class TestExtractor(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()
        self.excel_file = "test_registers.xlsx"
        self.pdf_file = "test_registers.pdf"
        self.mapping_file = "test_mapping.json"
        self.csv_file = "test_registers.csv"

        if HAS_OPENPYXL:
            # Create dummy Excel
            wb = Workbook()
            ws = wb.active
            ws.title = "Registers"
            ws.append(["Reg Addr", "Description", "Data Type", "Unit"])
            ws.append(["0x0001", "Voltage", "Uint16", "V"])
            ws.append(["0x0002", "Current", "Int32", "A"])
            ws.append(["40001", "Power", "Float32", "W"])
            wb.save(self.excel_file)

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

        # Create dummy CSV
        with open(self.csv_file, 'w', encoding='utf-8') as f:
            f.write("Address;Name;Type\n")
            f.write("2000;Frequency;U16\n")
            f.write("2001;Status;U16\n")

    def tearDown(self):
        for f in [self.excel_file, self.pdf_file, self.mapping_file, self.csv_file]:
            if os.path.exists(f):
                os.remove(f)

    @unittest.skipUnless(HAS_OPENPYXL, "openpyxl not installed")
    def test_extract_from_excel(self):
        data = self.extractor.extract_from_excel(self.excel_file)
        self.assertEqual(len(data), 1) # one sheet
        self.assertEqual(len(data[0]), 3) # 3 rows
        self.assertEqual(str(data[0][0]["Reg Addr"]), "0x0001")

    @unittest.skipUnless(HAS_OPENPYXL, "openpyxl not installed")
    def test_map_and_clean_excel(self):
        raw_data = self.extractor.extract_from_excel(self.excel_file)
        self.extractor.mapping = {
            "Address": "Reg Addr",
            "Name": "Description",
            "Type": "Data Type"
        }
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(len(mapped), 3)
        self.assertEqual(mapped[0]["Address"], "0x0001")
        self.assertEqual(mapped[0]["Name"], "Voltage")

    @unittest.skipUnless(HAS_PDFPLUMBER and HAS_REPORTLAB, "pdfplumber or reportlab not installed")
    def test_extract_from_pdf(self):
        data = self.extractor.extract_from_pdf(self.pdf_file)
        self.assertEqual(len(data), 1) # one table
        self.assertEqual(len(data[0]), 2) # 2 rows (header excluded by extractor)
        self.assertEqual(data[0][0]["Address"], "1000")
        self.assertEqual(data[0][0]["Name"], "Temp")

    def test_extract_from_csv(self):
        data = self.extractor.extract_from_csv(self.csv_file)
        self.assertEqual(len(data), 1)
        self.assertEqual(len(data[0]), 2)
        self.assertEqual(data[0][0]["Address"], "2000")

    def test_fuzzy_mapping(self):
        raw_data = [
            {"Register Address": "0x10", "Variable Name": "Test", "Data Type": "Uint16"}
        ]
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "0x10")
        self.assertEqual(mapped[0]["Name"], "Test")

    def test_fraction_parsing(self):
        raw_data = [
            {"Address": "100", "Name": "Var", "Factor": "1/10"}
        ]
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Factor"], "0.1")

if __name__ == "__main__":
    unittest.main()
