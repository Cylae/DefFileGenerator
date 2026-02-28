import unittest
import os
import csv
import json
from openpyxl import Workbook
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib.pagesizes import letter
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.security_utils import ensure_safe_path

class TestExtractor(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()
        self.excel_file = "test_registers.xlsx"
        self.pdf_file = "test_registers.pdf"
        self.mapping_file = "test_mapping.json"
        self.csv_file = "test_registers.csv"

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
        with open(self.csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(["Address", "Name", "Type"])
            writer.writerow(["2000", "Freq", "F32"])

    def tearDown(self):
        for f in [self.excel_file, self.pdf_file, self.mapping_file, self.csv_file]:
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

    def test_extract_from_csv(self):
        tables = self.extractor.extract_from_csv(self.csv_file)
        self.assertEqual(len(tables), 1)
        data = tables[0]
        self.assertEqual(len(data), 1)
        self.assertEqual(data[0]["Address"], "2000")

    def test_fuzzy_mapping(self):
        # Even without explicit mapping, it should find Name, Address, Type if headers are similar
        raw_data = [
            {"Register Address": "0x10", "Variable Name": "Test", "Data Type": "Uint16"}
        ]
        mapped = self.extractor.map_and_clean([raw_data])
        self.assertEqual(mapped[0]["Address"], "16")
        self.assertEqual(mapped[0]["Name"], "Test")
        self.assertEqual(mapped[0]["Type"], "U16")

class TestSecurity(unittest.TestCase):
    def test_ensure_safe_path(self):
        base = os.getcwd()
        # Safe paths
        self.assertTrue(ensure_safe_path("file.txt").startswith(base))
        self.assertTrue(ensure_safe_path("./file.txt").startswith(base))

        # Unsafe paths
        with self.assertRaises(ValueError):
            ensure_safe_path("../outside.txt")
        with self.assertRaises(ValueError):
            ensure_safe_path("/etc/passwd")

if __name__ == "__main__":
    unittest.main()
