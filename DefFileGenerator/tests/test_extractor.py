import unittest
import os
import csv
import json
import io

from DefFileGenerator.extractor import Extractor, HAS_OPENPYXL, HAS_PDFPLUMBER, HAS_DEFUSEDXML

class TestExtractor(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()
        self.excel_file = "test_registers.xlsx"
        self.pdf_file = "test_registers.pdf"
        self.mapping_file = "test_mapping.json"
        self.xml_file = "test_registers.xml"

        if HAS_OPENPYXL:
            from openpyxl import Workbook
            # Create dummy Excel
            wb = Workbook()
            ws = wb.active
            ws.title = "Registers"
            ws.append(["Reg Addr", "Description", "Data Type", "Unit"])
            ws.append(["0x0001", "Voltage", "Uint16", "V"])
            ws.append(["0x0002", "Current", "Int32", "A"])
            ws.append(["40001", "Power", "Float32", "W"])
            wb.save(self.excel_file)

        if HAS_PDFPLUMBER:
            try:
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
            except ImportError:
                pass

        # Create dummy XML
        with open(self.xml_file, "w") as f:
            f.write("<registers><register><Address>100</Address><Name>Test</Name></register></registers>")

    def tearDown(self):
        for f in [self.excel_file, self.pdf_file, self.mapping_file, self.xml_file]:
            if os.path.exists(f):
                os.remove(f)

    @unittest.skipUnless(HAS_OPENPYXL, "openpyxl not installed")
    def test_extract_from_excel(self):
        data = self.extractor.extract_from_excel(self.excel_file)
        self.assertEqual(len(data), 1)
        self.assertEqual(len(data[0]), 3)
        self.assertEqual(str(data[0][0]["Reg Addr"]), "0x0001")

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
        self.assertEqual(len(data), 1)
        self.assertEqual(len(data[0]), 2)
        self.assertEqual(data[0][0]["Address"], "1000")
        self.assertEqual(data[0][0]["Name"], "Temp")

    def test_fuzzy_mapping(self):
        # Even without explicit mapping, it should find Name, Address, Type if headers are similar
        raw_data = [
            {"Register Address": "0x10", "Variable Name": "Test", "Data Type": "Uint16"}
        ]
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "16")
        self.assertEqual(mapped[0]["Name"], "Test")
        self.assertEqual(mapped[0]["Type"], "U16")

    @unittest.skipUnless(HAS_DEFUSEDXML, "defusedxml not installed")
    def test_extract_from_xml(self):
        data = self.extractor.extract_from_xml(self.xml_file)
        self.assertEqual(len(data), 1)
        self.assertEqual(data[0][0]["Address"], "100")
        self.assertEqual(data[0][0]["Name"], "Test")

    @unittest.skipUnless(HAS_DEFUSEDXML, "defusedxml not installed")
    def test_xxe_protection(self):
        xxe_file = "test_xxe.xml"
        payload = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE root [
  <!ENTITY xxe SYSTEM "file:///etc/passwd">
]>
<root>
  <row>
    <Name>&xxe;</Name>
    <Address>100</Address>
  </row>
</root>"""
        with open(xxe_file, "w") as f:
            f.write(payload)
        try:
            from defusedxml import EntitiesForbidden
            with self.assertRaises(EntitiesForbidden):
                self.extractor.extract_from_xml(xxe_file)
        finally:
            if os.path.exists(xxe_file):
                os.remove(xxe_file)

    def test_compound_address_bits(self):
        raw_data = [
            {"Address": "100", "Name": "BitVar", "start": "0", "len": "1"}
        ]
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "100_0_1")
        self.assertEqual(mapped[0]["Type"], "BITS")

    def test_compound_address_string(self):
        raw_data = [
            {"Address": "200", "Name": "StrVar", "Type": "STRING", "len": "10"}
        ]
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "200_10")
        self.assertEqual(mapped[0]["Type"], "STRING")

if __name__ == "__main__":
    unittest.main()
