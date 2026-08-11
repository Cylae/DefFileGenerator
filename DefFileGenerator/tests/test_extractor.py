import unittest
import os
import csv
import json
import itertools

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

from DefFileGenerator.extractor import Extractor, peek_generator

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

    def test_peek_generator(self):
        gen = (i for i in range(3))
        has_data, new_gen = peek_generator(gen)
        self.assertTrue(has_data)
        self.assertEqual(list(new_gen), [0, 1, 2])

        empty_gen = (i for i in range(0))
        has_data, new_gen = peek_generator(empty_gen)
        self.assertFalse(has_data)
        self.assertEqual(list(new_gen), [])

    def test_normalize_type(self):
        self.assertEqual(Extractor.normalize_type("Uint16"), "U16")
        self.assertEqual(Extractor.normalize_type("Int32"), "I32")
        self.assertEqual(Extractor.normalize_type("Float32"), "F32")
        self.assertEqual(Extractor.normalize_type("unsigned int 16"), "U16")

    @unittest.skipUnless(HAS_OPENPYXL, "openpyxl not installed")
    def test_extract_from_excel(self):
        if not os.path.exists(self.excel_file): self.skipTest("Excel file not created")
        data = [list(table) for table in self.extractor.extract_from_excel(self.excel_file)]
        self.assertEqual(len(data), 1)
        self.assertEqual(len(data[0]), 3)
        self.assertEqual(str(data[0][0]["Reg Addr"]), "0x0001")

    @unittest.skipUnless(HAS_OPENPYXL, "openpyxl not installed")
    def test_map_and_clean_excel(self):
        if not os.path.exists(self.excel_file): self.skipTest("Excel file not created")
        raw_data = self.extractor.extract_from_excel(self.excel_file)
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

    def test_fuzzy_mapping(self):
        raw_data = [
            [{"Register Address": "0x10", "Variable Name": "Test", "Data Type": "Uint16"}]
        ]
        mapped = list(self.extractor.map_and_clean(raw_data))
        self.assertEqual(mapped[0]["Address"], "16")
        self.assertEqual(mapped[0]["Name"], "Test")
        self.assertEqual(mapped[0]["Type"], "U16")

    def test_pdf_page_validation(self):
        # Mocking pdfplumber.open is complex, but we can test the input handling
        self.extractor.extract_from_pdf("nonexistent.pdf", pages="1,2,3")
        # Should log error and return empty iterator (tested via coverage or manual run)

    def test_peek_generator_empty(self):
        from DefFileGenerator.def_gen import peek_generator
        has_data, it = peek_generator([])
        self.assertFalse(has_data)
        self.assertEqual(list(it), [])

    def test_peek_generator_with_data(self):
        from DefFileGenerator.def_gen import peek_generator
        has_data, it = peek_generator([1, 2, 3])
        self.assertTrue(has_data)
        self.assertEqual(list(it), [1, 2, 3])

    def test_process_row_edge_cases(self):
        raw_data = [[
            # Missing Name and Address -> should be filtered out
            {"Register Address": "", "Variable Name": "", "Data Type": "U16"},

            # BITS type with StartBit and Length -> appends _start_len
            {"Register Address": "100", "Variable Name": "Status", "Data Type": "BITS", "StartBit": "2", "Length": "4"},

            # BITS type with StartBit but missing Length -> defaults to length 1
            {"Register Address": "101", "Variable Name": "Flag", "Data Type": "BITS", "StartBit": "3", "Length": ""},

            # STRING type with Length -> appends _len
            {"Register Address": "200", "Variable Name": "Serial", "Data Type": "STRING", "Length": "10"},

            # STRING type without Length -> unchanged address
            {"Register Address": "201", "Variable Name": "Version", "Data Type": "STRING", "Length": ""},

            # Address already containing '_' -> does not append
            {"Register Address": "102_4_2", "Variable Name": "SubStatus", "Data Type": "BITS", "StartBit": "4", "Length": "2"},

            # Valid Factor field processing -> converts to string
            {"Register Address": "300", "Variable Name": "FactorTest", "Data Type": "U16", "ScaleFactor": "0.1"},

            # Missing/empty RegisterType -> falls back to 'Holding Register'
            {"Register Address": "400", "Variable Name": "RegTypeTest", "Data Type": "U16", "RegisterType": ""},

            # Correct application of address_offset
            {"Register Address": "500", "Variable Name": "OffsetTest", "Data Type": "U16"}
        ]]

        # We need an Extractor instance
        extractor = Extractor()

        # map_and_clean uses the fuzzy mapping, so we need to set the address offset
        mapped = list(extractor.map_and_clean(raw_data, address_offset=10))

        self.assertEqual(len(mapped), 8) # One row filtered out

        # BITS with StartBit and Length
        self.assertEqual(mapped[0]['Address'], '110_2_4')
        self.assertEqual(mapped[0]['Type'], 'BITS')

        # BITS with StartBit but missing Length
        self.assertEqual(mapped[1]['Address'], '111_3_1')
        self.assertEqual(mapped[1]['Type'], 'BITS')

        # STRING with Length
        self.assertEqual(mapped[2]['Address'], '210_10')
        self.assertEqual(mapped[2]['Type'], 'STRING')

        # STRING without Length
        self.assertEqual(mapped[3]['Address'], '211')
        self.assertEqual(mapped[3]['Type'], 'STRING')

        # Address already containing '_'
        self.assertEqual(mapped[4]['Address'], '112_4_2')
        self.assertEqual(mapped[4]['Type'], 'BITS')

        # Valid Factor field processing
        self.assertEqual(mapped[5]['Address'], '310')
        self.assertEqual(mapped[5]['ScaleFactor'], '0.1')

        # Missing/empty RegisterType
        self.assertEqual(mapped[6]['Address'], '410')
        self.assertEqual(mapped[6]['RegisterType'], 'Holding Register')

        # Correct application of address_offset
        self.assertEqual(mapped[7]['Address'], '510')

if __name__ == "__main__":
    unittest.main()
