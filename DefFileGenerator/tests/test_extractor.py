import os
import unittest

try:
    from openpyxl import Workbook  # noqa: F401

    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas  # noqa: F401
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

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
            data = [["Address", "Name", "Type"], ["1000", "Temp", "U16"], ["1001", "Humid", "U16"]]
            t = Table(data)
            t.setStyle(
                TableStyle(
                    [
                        ("GRID", (0, 0), (-1, -1), 1, (0, 0, 0)),
                    ]
                )
            )
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
        self.assertEqual(Extractor.normalize_type("U32 swap"), "U32_WB")
        self.assertEqual(Extractor.normalize_type("I16_B"), "I16_B")
        self.assertEqual(Extractor.normalize_type("F64 big endian"), "F64_WB")

    @unittest.skipUnless(HAS_OPENPYXL, "openpyxl not installed")
    def test_extract_from_excel(self):
        if not os.path.exists(self.excel_file):
            self.skipTest("Excel file not created")
        data = [list(table) for table in self.extractor.extract_from_excel(self.excel_file)]
        self.assertEqual(len(data), 1)
        self.assertEqual(len(data[0]), 3)
        self.assertEqual(str(data[0][0]["Reg Addr"]), "0x0001")

    @unittest.skipUnless(HAS_OPENPYXL, "openpyxl not installed")
    def test_map_and_clean_excel(self):
        if not os.path.exists(self.excel_file):
            self.skipTest("Excel file not created")
        raw_data = self.extractor.extract_from_excel(self.excel_file)
        self.extractor.mapping = {"Address": "Reg Addr", "Name": "Description", "Type": "Data Type"}
        mapped = list(self.extractor.map_and_clean(raw_data))
        self.assertEqual(len(mapped), 3)
        self.assertEqual(mapped[0]["Address"], "1")
        self.assertEqual(mapped[0]["Name"], "Voltage")
        self.assertEqual(mapped[0]["Type"], "U16")

    def test_fuzzy_mapping(self):
        raw_data = [[{"Register Address": "0x10", "Variable Name": "Test", "Data Type": "Uint16"}]]
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
        # We test the internal `process_row` by passing edge-case rows through `map_and_clean`
        # Extractor.map_and_clean expects an iterable of iterables (tables)
        # We will create one table with various rows that trigger edge cases.
        raw_data = [
            [
                # 1. Missing Name and Address -> Skipped
                {"Data Type": "U16"},
                # 2. Missing Type -> Defaults to U16, missing RegisterType -> Holding Register
                {"Address": "100", "Name": "Default Type Test"},
                # 3. BITS type with StartBit and Length
                {
                    "Address": "101",
                    "Name": "Bits Test 1",
                    "Type": "BITS",
                    "StartBit": "2",
                    "Length": "4",
                },
                # 4. BITS type with StartBit only
                {"Address": "102", "Name": "Bits Test 2", "Type": "BITS", "StartBit": "3"},
                # 5. STRING type with Length
                {"Address": "103", "Name": "String Test", "Type": "STRING", "Length": "10"},
                # 6. Pre-existing underscore in Address -> No suffix appended
                {
                    "Address": "104_1_2",
                    "Name": "Underscore Bits Test",
                    "Type": "BITS",
                    "StartBit": "3",
                    "Length": "4",
                },
                {
                    "Address": "105_10",
                    "Name": "Underscore String Test",
                    "Type": "STRING",
                    "Length": "20",
                },
                # 7. Factor normalization
                {"Address": "106", "Name": "Factor Test", "Factor": "0.5"},
            ]
        ]

        # We need to set up Extractor mapping so it finds the columns properly
        # Or rely on standard column detection which is already good enough for Address, Name, Type, StartBit, Length, Factor
        extractor = Extractor()

        # We need an explicit mapping or standard columns:
        # The standard column mappings should pick up "Address", "Name", "Type", "StartBit", "Length", "Factor"
        mapped = list(extractor.map_and_clean(raw_data))

        # Check that row 1 (missing name/addr) was skipped
        # So we expect exactly 7 mapped rows in total
        self.assertEqual(len(mapped), 7)

        # 2. Default Type and Default RegisterType
        self.assertEqual(mapped[0]["Address"], "100")
        self.assertEqual(mapped[0]["Type"], "U16")
        self.assertEqual(mapped[0]["RegisterType"], "Holding Register")

        # 3. BITS type with StartBit and Length
        self.assertEqual(mapped[1]["Address"], "101_2_4")
        self.assertEqual(mapped[1]["Type"], "BITS")

        # 4. BITS type with StartBit only -> Length defaults to 1
        self.assertEqual(mapped[2]["Address"], "102_3_1")

        # 5. STRING type with Length
        self.assertEqual(mapped[3]["Address"], "103_10")
        self.assertEqual(mapped[3]["Type"], "STRING")

        # 6. Pre-existing underscore in Address
        self.assertEqual(mapped[4]["Address"], "104_1_2")
        self.assertEqual(mapped[5]["Address"], "105_10")

        # 7. Factor normalization -> 0.5 is parsed as float then string
        self.assertEqual(mapped[6]["Address"], "106")
        self.assertEqual(mapped[6]["Factor"], "0.5")

    def test_gain_column_is_inverted_into_factor(self):
        # Manufacturer tables (Huawei-style) express a divisor "Gain" rather than
        # a direct multiplier. Factor (CoefA) must be its reciprocal.
        raw_data = [
            [
                {"Address": "30073", "Name": "Rated power (Pn)", "Type": "U32", "Gain": "1000"},
            ]
        ]
        mapped = list(self.extractor.map_and_clean(raw_data))
        self.assertEqual(len(mapped), 1)
        self.assertAlmostEqual(float(mapped[0]["Factor"]), 0.001)

    def test_explicit_factor_column_takes_precedence_over_gain(self):
        raw_data = [
            [
                {"Address": "1", "Name": "Test", "Type": "U16", "Factor": "2", "Gain": "1000"},
            ]
        ]
        mapped = list(self.extractor.map_and_clean(raw_data))
        self.assertEqual(float(mapped[0]["Factor"]), 2.0)

    def test_read_write_column_maps_to_action(self):
        raw_data = [
            [
                {"Address": "30070", "Name": "Model ID", "Type": "U16", "Read/ Write": "RO"},
                {"Address": "47100", "Name": "Setting", "Type": "U16", "Read/ Write": "RW"},
            ]
        ]
        mapped = list(self.extractor.map_and_clean(raw_data))
        self.assertEqual(mapped[0]["Action"], "RO")
        self.assertEqual(mapped[1]["Action"], "RW")

    def test_string_quantity_column_is_doubled_to_byte_length(self):
        # A "Quantity" header expresses a register (16-bit word) count; the
        # WebdynSunPM STRING address suffix must be a byte length (2x).
        raw_data = [
            [
                {"Address": "30000", "Name": "Model", "Type": "STR", "Quantity": "15"},
            ]
        ]
        mapped = list(self.extractor.map_and_clean(raw_data))
        self.assertEqual(mapped[0]["Address"], "30000_30")
        self.assertEqual(mapped[0]["Type"], "STRING")

    def test_length_column_is_not_doubled(self):
        # A plain "Length" header is already assumed to be a byte length and
        # must NOT be doubled (distinct from "Quantity"/"Count").
        raw_data = [
            [
                {"Address": "103", "Name": "String Test", "Type": "STRING", "Length": "10"},
            ]
        ]
        mapped = list(self.extractor.map_and_clean(raw_data))
        self.assertEqual(mapped[0]["Address"], "103_10")


if __name__ == "__main__":
    unittest.main()


class TestExtractorUncoveredEdgeCases(unittest.TestCase):
    def test_extract_excel_missing_dependency(self):
        extractor = Extractor()
        import DefFileGenerator.extractor as ext_mod

        old_val = ext_mod.HAS_OPENPYXL
        try:
            ext_mod.HAS_OPENPYXL = False
            tables = list(extractor.extract_from_excel("dummy.xlsx"))
            self.assertEqual(len(tables), 0)
        finally:
            ext_mod.HAS_OPENPYXL = old_val

    def test_extract_pdf_missing_dependency(self):
        extractor = Extractor()
        import DefFileGenerator.extractor as ext_mod

        old_val = ext_mod.HAS_PDFPLUMBER
        try:
            ext_mod.HAS_PDFPLUMBER = False
            tables = list(extractor.extract_from_pdf("dummy.pdf"))
            self.assertEqual(len(tables), 0)
        finally:
            ext_mod.HAS_PDFPLUMBER = old_val

    def test_extract_xml_missing_dependency(self):
        extractor = Extractor()
        import DefFileGenerator.extractor as ext_mod

        old_val = ext_mod.HAS_DEFUSEDXML
        try:
            ext_mod.HAS_DEFUSEDXML = False
            tables = list(extractor.extract_from_xml("dummy.xml"))
            self.assertEqual(len(tables), 0)
        finally:
            ext_mod.HAS_DEFUSEDXML = old_val
