import unittest
from unittest.mock import MagicMock, patch
import logging
from DefFileGenerator.extractor import Extractor

class TestPDFValidation(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()

    @patch('pdfplumber.open')
    def test_extract_from_pdf_page_validation(self, mock_pdf_open):
        # Setup mock PDF with 5 pages
        mock_pdf = MagicMock()
        mock_pdf.pages = [MagicMock(page_number=i) for i in range(1, 6)]
        for p in mock_pdf.pages:
            p.extract_tables.return_value = [[['Header'], ['Row']]] # One table per page
        mock_pdf_open.return_value.__enter__.return_value = mock_pdf

        # Test valid pages - no warnings expected
        # We don't use assertLogs here because it fails if no logs are emitted
        gen = self.extractor.extract_from_pdf('dummy.pdf', pages=[1, 3, 5])
        list(gen)

        # Test out of range pages
        with self.assertLogs(level='WARNING') as cm:
             gen = self.extractor.extract_from_pdf('dummy.pdf', pages=[0, 6, 2])
             list(gen)

        self.assertTrue(any("Page 0 is out of range" in log for log in cm.output))
        self.assertTrue(any("Page 6 is out of range" in log for log in cm.output))

    @patch('pdfplumber.open')
    def test_extract_from_pdf_all_pages(self, mock_pdf_open):
        mock_pdf = MagicMock()
        mock_pdf.pages = [MagicMock(page_number=1), MagicMock(page_number=2)]
        for p in mock_pdf.pages:
            p.extract_tables.return_value = [[['Header'], ['Row']]]
        mock_pdf_open.return_value.__enter__.return_value = mock_pdf

        # pages=None should extract from all pages
        gen = self.extractor.extract_from_pdf('dummy.pdf', pages=None)
        list(gen)

        self.assertEqual(mock_pdf.pages[0].extract_tables.call_count, 1)
        self.assertEqual(mock_pdf.pages[1].extract_tables.call_count, 1)

if __name__ == '__main__':
    unittest.main()
