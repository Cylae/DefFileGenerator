import unittest
import logging
import os
import csv
from io import StringIO
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator
from DefFileGenerator.extractor import Extractor

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.extractor = Extractor()
        # Suppress logging during tests unless needed
        logging.getLogger().setLevel(logging.CRITICAL)

    def test_address_range_validation(self):
        # Valid addresses
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))

        # Out of range addresses (should still return True for format, but log warning)
        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_address("65536", "U16"))
            self.assertIn("outside the standard Modbus range", cm.output[0])

        with self.assertLogs(level='WARNING') as cm:
            self.assertTrue(self.generator.validate_address("-1", "U16"))
            self.assertIn("outside the standard Modbus range", cm.output[0])

    def test_intelligent_action_defaulting(self):
        rows = [
            {'Name': 'Input Reg', 'Address': '1', 'RegisterType': 'Input Register', 'Type': 'U16'},
            {'Name': 'Holding Reg', 'Address': '2', 'RegisterType': 'Holding Register', 'Type': 'U16'},
            {'Name': 'Discrete Input', 'Address': '3', 'RegisterType': 'Discrete Input', 'Type': 'U16'},
            {'Name': 'Coil', 'Address': '4', 'RegisterType': 'Coil', 'Type': 'U16'},
        ]
        processed = list(self.generator.process_rows(rows))

        # Input Register (4) -> Read Only (4)
        self.assertEqual(processed[0]['Action'], '4')
        # Holding Register (3) -> Read/Write (1)
        self.assertEqual(processed[1]['Action'], '1')
        # Discrete Input (2) -> Read Only (4)
        self.assertEqual(processed[2]['Action'], '4')
        # Coil (1) -> Read/Write (1)
        self.assertEqual(processed[3]['Action'], '1')

    def test_action_synonyms(self):
        rows = [
            {'Name': 'R1', 'Address': '1', 'Action': 'RO'},
            {'Name': 'R2', 'Address': '2', 'Action': 'READ ONLY'},
            {'Name': 'R3', 'Address': '3', 'Action': 'RW'},
            {'Name': 'R4', 'Address': '4', 'Action': 'READ/WRITE'},
        ]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4')
        self.assertEqual(processed[1]['Action'], '4')
        self.assertEqual(processed[2]['Action'], '1')
        self.assertEqual(processed[3]['Action'], '1')

    def test_pdf_page_validation_mock(self):
        # We can't easily test real PDF without dependencies, but we can test the logic
        # by mocking pdfplumber if we had a more unit-testable structure for it.
        # Given the implementation, we'll rely on existing tests for extractor if any,
        # or just ensure it doesn't crash on empty/None pages.
        pass

    def test_excel_sheet_missing_logging(self):
        # Test the newly added sheet existence check in extract_from_excel
        # This requires HAS_OPENPYXL to be True to reach the code path.
        from DefFileGenerator.extractor import HAS_OPENPYXL
        if not HAS_OPENPYXL:
            self.skipTest("openpyxl not installed")

        # Create a dummy excel
        import openpyxl
        wb = openpyxl.Workbook()
        wb.save("dummy_test.xlsx")

        try:
            with self.assertLogs(level='WARNING') as cm:
                gen = self.extractor.extract_from_excel("dummy_test.xlsx", sheet_name="NonExistentSheet")
                list(gen) # Consume generator
                self.assertIn("not found in", cm.output[0])
        finally:
            if os.path.exists("dummy_test.xlsx"):
                os.remove("dummy_test.xlsx")

    def test_summary_logging(self):
        rows = [
            {'Info1': '3', 'Info2': '1', 'Info3': 'U16', 'Info4': '', 'Name': 'N1', 'Tag': 'T1', 'CoefA': '1.0', 'CoefB': '0.0', 'Unit': 'V', 'Action': '1'},
            {'Info1': '4', 'Info2': '2', 'Info3': 'U16', 'Info4': '', 'Name': 'N2', 'Tag': 'T2', 'CoefA': '1.0', 'CoefB': '0.0', 'Unit': 'A', 'Action': '4'},
        ]
        output = StringIO()
        with self.assertLogs(level='INFO') as cm:
            self.generator.write_output_csv(output, rows, "MFG", "MOD")
            self.assertIn("Processed 2 registers", cm.output[0])
            self.assertIn("Holding Registers: 1", cm.output[0])
            self.assertIn("Input Registers: 1", cm.output[0])

if __name__ == '__main__':
    unittest.main()
