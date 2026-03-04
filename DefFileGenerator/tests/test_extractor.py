import unittest
import os
import csv
import json
import logging

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
from DefFileGenerator.def_gen import Generator

class TestExtractor(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()
        self.excel_file = "test_registers.xlsx"
        self.pdf_file = "test_registers.pdf"
        self.mapping_file = "test_mapping.json"
        self.csv_file = "test_csv_output.csv"

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
        for f in [self.excel_file, self.pdf_file, self.mapping_file, self.csv_file]:
            if os.path.exists(f):
                os.remove(f)

    def test_normalize_action(self):
        self.assertEqual(self.extractor.normalize_action("R"), "4")
        self.assertEqual(self.extractor.normalize_action("RW"), "1")
        self.assertEqual(self.extractor.normalize_action("read"), "4")
        self.assertEqual(self.extractor.normalize_action("WRITE"), "1")

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
        self.assertEqual(mapped[0]["Address"], "1")
        self.assertEqual(mapped[0]["Name"], "Voltage")

        # Test type normalization (delegated to Generator by Generator.process_rows)
        # In map_and_clean, Type is not normalized unless it's STRING.
        # But we can verify it was preserved correctly
        self.assertEqual(mapped[0]["Type"], "Uint16")

        # Now test with Generator to see it normalizes correctly
        gen = Generator()
        processed = gen.process_rows(mapped)
        self.assertEqual(processed[0]["Info3"], "U16")
        self.assertEqual(processed[1]["Info3"], "I32")
        self.assertEqual(processed[2]["Info3"], "F32")

    @unittest.skipUnless(HAS_PDFPLUMBER and HAS_REPORTLAB, "pdfplumber or reportlab not installed")
    def test_extract_from_pdf(self):
        tables = self.extractor.extract_from_pdf(self.pdf_file)
        self.assertEqual(len(tables), 1)
        data = tables[0]
        self.assertEqual(len(data), 2)
        self.assertEqual(data[0]["Address"], "1000")
        self.assertEqual(data[0]["Name"], "Temp")

    def test_extract_from_csv(self):
        with open(self.csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(["Address", "Name", "Type"])
            writer.writerow(["2000", "Volt", "U16"])
            writer.writerow(["2001", "Curr", "U16"])

        tables = self.extractor.extract_from_csv(self.csv_file)
        self.assertEqual(len(tables), 1)
        self.assertEqual(len(tables[0]), 2)
        self.assertEqual(tables[0][0]["Address"], "2000")

    def test_fuzzy_mapping(self):
        # Even without explicit mapping, it should find Name, Address, Type if headers are similar
        raw_data = [
            [{"Register Address": "0x10", "Variable Name": "Test", "Data Type": "Uint16"}]
        ]
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "16")
        self.assertEqual(mapped[0]["Name"], "Test")
        self.assertEqual(mapped[0]["Type"], "Uint16")

        # Verify Generator normalizes it
        gen = Generator()
        processed = gen.process_rows(mapped)
        self.assertEqual(processed[0]["Info3"], "U16")

if __name__ == "__main__":
    unittest.main()
