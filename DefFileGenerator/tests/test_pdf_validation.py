import unittest
from unittest.mock import patch, MagicMock
import logging
import io
from DefFileGenerator.extractor import Extractor

class TestPdfValidation(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()

    @patch('pdfplumber.open')
    def test_extract_from_pdf_page_range(self, mock_open):
        # Setup mock PDF
        mock_pdf = MagicMock()
        mock_pdf.pages = [MagicMock(), MagicMock(), MagicMock()] # 3 pages
        for i, page in enumerate(mock_pdf.pages):
            page.page_number = i + 1
            page.extract_tables.return_value = [[["Address", "Name"], ["100", "Test"]]]

        mock_open.return_value.__enter__.return_value = mock_pdf

        # Test valid page
        results = list(self.extractor.extract_from_pdf("dummy.pdf", pages=1))
        self.assertEqual(len(results), 1)

        # Test out of range page
        with self.assertLogs(level='WARNING') as cm:
            results = list(self.extractor.extract_from_pdf("dummy.pdf", pages=5))
            self.assertEqual(len(results), 0)
            self.assertTrue(any("Page 5 is out of range" in output for output in cm.output))

        # Test mixed valid/invalid pages
        with self.assertLogs(level='WARNING') as cm:
            results = list(self.extractor.extract_from_pdf("dummy.pdf", pages=[1, 10]))
            self.assertEqual(len(results), 1)
            self.assertTrue(any("Page 10 is out of range" in output for output in cm.output))

if __name__ == '__main__':
    unittest.main()
