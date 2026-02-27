import unittest
import os
import csv
import json
from openpyxl import Workbook
from reportlab.pdfgen import canvas
from DefFileGenerator.extractor import Extractor

class TestExtractor(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()
        self.excel_file = "test_registers.xlsx"
        self.pdf_file = "test_registers.pdf"
        self.mapping_file = "test_mapping.json"

        # Create dummy Excel
        wb = Workbook()
        ws = wb.active
        ws.title = "Registers"
        ws.append(["Reg Addr", "Description", "Data Type", "Unit"])
        ws.append(["0x0001", "Voltage", "Uint16", "V"])
        ws.append(["0x0002", "Current", "Int32", "A"])
        ws.append(["40001", "Power", "Float32", "W"])
        wb.save(self.excel_file)

        # Create dummy PDF
        c = canvas.Canvas(self.pdf_file)
        c.drawString(100, 800, "Register Map")
        # Simple table-like text (Note: pdfplumber works best with actual PDF tables,
        # but reportlab can create them if we use Table objects. For simplicity,
        # I'll just use the Excel one as primary and a simple PDF if I can)
        # Actually, creating a real table in PDF with reportlab is better for pdfplumber
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
        from reportlab.lib.pagesizes import letter

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
        self.assertEqual(self.extractor.normalize_type("Uint16"), "U16")
        self.assertEqual(self.extractor.normalize_type("Int32"), "I32")
        self.assertEqual(self.extractor.normalize_type("Float32"), "F32")
        self.assertEqual(self.extractor.normalize_type("unsigned int 16"), "U16")

    def test_extract_from_excel(self):
        tables = self.extractor.extract_from_excel(self.excel_file)
        self.assertEqual(len(tables), 1)
        data = tables[0]
        self.assertEqual(len(data), 3)
        self.assertEqual(str(data[0]["Reg Addr"]), "0x0001")

    def test_map_and_clean_excel(self):
        raw_tables = self.extractor.extract_from_excel(self.excel_file)
        # Custom mapping
        self.extractor.mapping = {
            "Address": "Reg Addr",
            "Name": "Description",
            "Type": "Data Type"
        }
        mapped = self.extractor.map_and_clean(raw_tables)
        self.assertEqual(len(mapped), 3)
        self.assertEqual(mapped[0]["Address"], "1")
        self.assertEqual(mapped[0]["Name"], "Voltage")
        self.assertEqual(mapped[0]["Type"], "U16")
        self.assertEqual(mapped[1]["Type"], "I32")
        self.assertEqual(mapped[2]["Type"], "F32")

    def test_extract_from_pdf(self):
        tables = self.extractor.extract_from_pdf(self.pdf_file)
        self.assertEqual(len(tables), 1)
        data = tables[0]
        self.assertEqual(len(data), 2)
        self.assertEqual(data[0]["Address"], "1000")
        self.assertEqual(data[0]["Name"], "Temp")

    def test_fuzzy_mapping(self):
        # Even without explicit mapping, it should find Name, Address, Type if headers are similar
        raw_tables = [[
            {"Register Address": "0x10", "Variable Name": "Test", "Data Type": "Uint16"}
        ]]
        mapped = self.extractor.map_and_clean(raw_tables)
        self.assertEqual(mapped[0]["Address"], "16")
        self.assertEqual(mapped[0]["Name"], "Test")
        self.assertEqual(mapped[0]["Type"], "U16")

    def test_extract_from_csv(self):
        csv_file = "test_registers.csv"
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(["Name", "Address", "Type"])
            writer.writerow(["Test1", "100", "U16"])

        try:
            tables = self.extractor.extract_from_csv(csv_file)
            self.assertEqual(len(tables), 1)
            self.assertEqual(len(tables[0]), 1)
            self.assertEqual(tables[0][0]["Name"], "Test1")
        finally:
            if os.path.exists(csv_file):
                os.remove(csv_file)

    def test_extract_from_xml(self):
        xml_file = "test_registers.xml"
        content = """<root>
            <row><Name>TestXML</Name><Address>200</Address><Type>U16</Type></row>
        </root>"""
        with open(xml_file, 'w') as f:
            f.write(content)

        try:
            tables = self.extractor.extract_from_xml(xml_file)
            self.assertEqual(len(tables), 1)
            self.assertEqual(len(tables[0]), 1)
            self.assertEqual(tables[0][0]["Name"], "TestXML")
        finally:
            if os.path.exists(xml_file):
                os.remove(xml_file)

if __name__ == "__main__":
    unittest.main()
