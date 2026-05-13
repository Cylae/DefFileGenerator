import unittest
from unittest.mock import patch, MagicMock
from DefFileGenerator.extractor import Extractor, HAS_PDFPLUMBER
import logging

class TestPDFValidation(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()

    @unittest.skipUnless(HAS_PDFPLUMBER, "pdfplumber not installed")
    @patch('pdfplumber.open')
    def test_pdf_page_range_validation(self, mock_open):
        # Mock PDF with 5 pages
        mock_pdf = MagicMock()
        mock_pdf.pages = [MagicMock() for _ in range(5)]
        for i, p in enumerate(mock_pdf.pages):
            p.page_number = i + 1
            p.extract_tables.return_value = []
        mock_open.return_value.__enter__.return_value = mock_pdf

        # Test valid and invalid pages
        with self.assertLogs(level='WARNING') as cm:
            # Page 1, 3 are valid. Page 10 is invalid.
            list(self.extractor.extract_from_pdf("dummy.pdf", pages=[1, 3, 10]))

            self.assertTrue(any("Page 10 requested but PDF has 5 pages" in output for output in cm.output))

    @patch('DefFileGenerator.extractor.HAS_PDFPLUMBER', False)
    def test_extract_from_pdf_no_dependency(self):
        with self.assertLogs(level='ERROR') as cm:
            gen = self.extractor.extract_from_pdf("dummy.pdf")
            result = list(gen)
            self.assertEqual(result, [])
            self.assertTrue(any("pdfplumber is required" in output for output in cm.output))

if __name__ == '__main__':
    unittest.main()
