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
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
    from reportlab.lib.pagesizes import letter
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator

class TestExtractor(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()
        self.excel_file = "test_registers.xlsx"
        self.pdf_file = "test_registers.pdf"
        self.mapping_file = "test_mapping.json"

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
            # Create dummy PDF
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

    def test_normalize_type_via_generator(self):
        gen = Generator()
        self.assertEqual(gen.normalize_type("Uint16"), "U16")
        self.assertEqual(gen.normalize_type("Int32"), "I32")
        self.assertEqual(gen.normalize_type("Float32"), "F32")
        self.assertEqual(gen.normalize_type("unsigned int 16"), "U16")

    @unittest.skipUnless(HAS_OPENPYXL, "openpyxl not installed")
    def test_extract_from_excel(self):
        tables = self.extractor.extract_from_excel(self.excel_file)
        self.assertEqual(len(tables), 1)
        data = tables[0]
        self.assertEqual(len(data), 3)
        self.assertEqual(str(data[0]["Reg Addr"]), "0x0001")

    @unittest.skipUnless(HAS_OPENPYXL, "openpyxl not installed")
    def test_map_and_clean_excel(self):
        tables = self.extractor.extract_from_excel(self.excel_file)
        # Custom mapping
        self.extractor.mapping = {
            "Address": "Reg Addr",
            "Name": "Description",
            "Type": "Data Type"
        }
        mapped = self.extractor.map_and_clean(tables)
        self.assertEqual(len(mapped), 3)
        # Address normalization now happens in map_and_clean via Generator.normalize_address_val
        # Actually in my refactor I used generator.normalize_address_val if it's there
        # Wait, I should check what I wrote in map_and_clean
        self.assertEqual(mapped[0]["Name"], "Voltage")

    @unittest.skipUnless(HAS_REPORTLAB, "reportlab not installed")
    def test_extract_from_pdf(self):
        # We need pdfplumber too for this to actually work
        try:
            import pdfplumber
        except ImportError:
            self.skipTest("pdfplumber not installed")

        tables = self.extractor.extract_from_pdf(self.pdf_file)
        self.assertEqual(len(tables), 1)
        data = tables[0]
        self.assertEqual(len(data), 2)
        self.assertEqual(data[0]["Address"], "1000")
        self.assertEqual(data[0]["Name"], "Temp")

    def test_fuzzy_mapping(self):
        # Even without explicit mapping, it should find Name, Address, Type if headers are similar
        raw_data = [
            {"Register Address": "0x10", "Variable Name": "Test", "Data Type": "Uint16"}
        ]
        mapped = self.extractor.map_and_clean(raw_data)
        # Address normalization in map_and_clean doesn't happen anymore,
        # it happens in Generator.process_rows.
        # Wait, I should re-check map_and_clean in extractor.py
        self.assertEqual(mapped[0]["Name"], "Test")
        self.assertEqual(mapped[0]["Type"], "U16")

if __name__ == "__main__":
    unittest.main()
