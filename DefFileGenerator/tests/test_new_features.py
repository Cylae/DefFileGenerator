import unittest
import logging
from unittest.mock import patch, MagicMock
from DefFileGenerator.def_gen import Generator, peek_generator
from DefFileGenerator.extractor import Extractor

class TestNewFeatures(unittest.TestCase):
    def test_peek_generator(self):
        # Empty
        res, it = peek_generator([])
        self.assertFalse(res)
        self.assertEqual(list(it), [])

        # Non-empty
        res, it = peek_generator([1, 2, 3])
        self.assertTrue(res)
        self.assertEqual(list(it), [1, 2, 3])

        # None
        res, it = peek_generator(None)
        self.assertFalse(res)
        self.assertEqual(list(it), [])

    def test_address_range_validation(self):
        gen = Generator()
        # Valid range - should NOT log
        res = gen.validate_address("0", "U16")
        self.assertTrue(res)

        res = gen.validate_address("65535", "U16")
        self.assertTrue(res)

        # Invalid range (outside 0-65535) - SHOULD log
        with self.assertLogs(level='WARNING') as cm:
             res = gen.validate_address("65536", "U16")
             self.assertTrue(res) # Format is still valid
             self.assertTrue(any("outside standard Modbus range" in r.getMessage() for r in cm.records))

        with self.assertLogs(level='WARNING') as cm:
             res = gen.validate_address("-1", "U16")
             self.assertTrue(res) # Format is still valid
             self.assertTrue(any("outside standard Modbus range" in r.getMessage() for r in cm.records))

    def test_intelligent_action_defaulting(self):
        gen = Generator()
        # Input Register -> Action 4
        rows = [{'Name': 'Test', 'Address': '1', 'RegisterType': 'Input Register', 'Type': 'U16'}]
        processed = list(gen.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4')

        # Discrete Input -> Action 4
        rows = [{'Name': 'Test', 'Address': '1', 'RegisterType': 'Discrete Input', 'Type': 'U16'}]
        processed = list(gen.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4')

        # Holding Register -> Action 1
        rows = [{'Name': 'Test', 'Address': '1', 'RegisterType': 'Holding Register', 'Type': 'U16'}]
        processed = list(gen.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

        # Coil -> Action 1
        rows = [{'Name': 'Test', 'Address': '1', 'RegisterType': 'Coil', 'Type': 'U16'}]
        processed = list(gen.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

    @patch('pdfplumber.open')
    def test_pdf_page_validation(self, mock_pdf_open):
        mock_pdf = MagicMock()
        mock_page = MagicMock()
        mock_page.page_number = 1
        mock_page.extract_tables.return_value = [[['Header'], ['Data']]]
        mock_pdf.pages = [mock_page, MagicMock(), MagicMock()] # 3 pages
        mock_pdf_open.return_value.__enter__.return_value = mock_pdf

        ext = Extractor()

        # Valid page
        gen = ext.extract_from_pdf("dummy.pdf", pages=[1])
        tables = list(gen)
        self.assertEqual(len(tables), 1)

        # Out of range page
        with self.assertLogs(level='WARNING') as cm:
            gen = ext.extract_from_pdf("dummy.pdf", pages=[5])
            list(gen)
            self.assertTrue(any("Page 5 is out of range" in r.getMessage() for r in cm.records))

        # String pages input
        gen = ext.extract_from_pdf("dummy.pdf", pages="1,2")
        tables = list(gen)
        self.assertTrue(len(tables) >= 1)

    def test_excel_sheet_missing(self):
        with patch('openpyxl.load_workbook') as mock_load:
            mock_wb = MagicMock()
            mock_wb.sheetnames = ['Sheet1']
            mock_load.return_value = mock_wb

            ext = Extractor()
            with self.assertLogs(level='WARNING') as cm:
                res = list(ext.extract_from_excel("dummy.xlsx", sheet_name="NonExistent"))
                self.assertEqual(res, [])
                self.assertTrue(any("Sheet 'NonExistent' not found" in r.getMessage() for r in cm.records))

if __name__ == '__main__':
    unittest.main()
