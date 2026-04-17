import unittest
import os
import csv
import json
from DefFileGenerator.extractor import Extractor, HAS_OPENPYXL, HAS_PDFPLUMBER, HAS_DEFUSEDXML

class TestExtractor(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()
        self.excel_file = "test_registers.xlsx"
        self.pdf_file = "test_registers.pdf"
        self.mapping_file = "test_mapping.json"
        self.csv_file = "test_registers.csv"

        if HAS_OPENPYXL:
            from openpyxl import Workbook
            wb = Workbook()
            ws = wb.active
            ws.title = "Registers"
            ws.append(["Reg Addr", "Description", "Data Type", "Unit"])
            ws.append(["0x0001", "Voltage", "Uint16", "V"])
            ws.append(["0x0002", "Current", "Int32", "A"])
            ws.append(["40001", "Power", "Float32", "W"])
            wb.save(self.excel_file)

        if HAS_PDFPLUMBER:
            from reportlab.pdfgen import canvas
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

        with open(self.csv_file, 'w', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=["Address", "Name", "Type", "Start", "Len"])
            writer.writeheader()
            writer.writerow({"Address": "100", "Name": "Var1", "Type": "U16", "Start": "0", "Len": "1"})
            writer.writerow({"Address": "200", "Name": "Var2", "Type": "BITS", "Start": "0", "Len": "8"})

    def tearDown(self):
        for f in [self.excel_file, self.pdf_file, self.mapping_file, self.csv_file]:
            if os.path.exists(f):
                os.remove(f)

    @unittest.skipUnless(HAS_OPENPYXL, "openpyxl not installed")
    def test_extract_from_excel(self):
        data = self.extractor.extract_from_excel(self.excel_file)
        self.assertEqual(len(data), 3)
        self.assertEqual(str(data[0]["Reg Addr"]), "0x0001")

    @unittest.skipUnless(HAS_OPENPYXL, "openpyxl not installed")
    def test_map_and_clean_excel(self):
        raw_data = self.extractor.extract_from_excel(self.excel_file)
        # Custom mapping
        self.extractor.mapping = {
            "Address": "Reg Addr",
            "Name": "Description",
            "Type": "Data Type"
        }
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(len(mapped), 3)
        self.assertEqual(mapped[0]["Address"], "1")
        self.assertEqual(mapped[0]["Name"], "Voltage")
        self.assertEqual(mapped[0]["Type"], "U16")
        self.assertEqual(mapped[1]["Type"], "I32")
        self.assertEqual(mapped[2]["Type"], "F32")

    @unittest.skipUnless(HAS_PDFPLUMBER, "pdfplumber not installed")
    def test_extract_from_pdf(self):
        data = self.extractor.extract_from_pdf(self.pdf_file)
        self.assertEqual(len(data), 2)
        self.assertEqual(data[0]["Address"], "1000")
        self.assertEqual(data[0]["Name"], "Temp")

    def test_fuzzy_mapping(self):
        # Even without explicit mapping, it should find Name, Address, Type if headers are similar
        raw_data = [
            {"Register Address": "0x10", "Variable Name": "Test", "Data Type": "Uint16"}
        ]
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "16")
        self.assertEqual(mapped[0]["Name"], "Test")
        self.assertEqual(mapped[0]["Type"], "U16")

    def test_separate_bit_len_columns(self):
        raw_data = [
            {"Address": "100", "Name": "BitVar", "Type": "BITS", "Start": "0", "Len": "1"},
            {"Address": "101", "Name": "BitVar2", "Type": "BITS", "Start": "4"} # Default len 1
        ]
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "100_0_1")
        self.assertEqual(mapped[1]["Address"], "101_4_1")

    @unittest.skipUnless(HAS_DEFUSEDXML, "defusedxml not installed")
    def test_extract_from_xml(self):
        xml_file = "test.xml"
        with open(xml_file, "w") as f:
            f.write("<root><reg><Address>100</Address><Name>Var1</Name></reg></root>")
        try:
            data = self.extractor.extract_from_xml(xml_file)
            self.assertEqual(len(data), 1)
            self.assertEqual(data[0]["Address"], "100")
        finally:
            if os.path.exists(xml_file):
                os.remove(xml_file)

if __name__ == "__main__":
    unittest.main()
