import unittest
import logging
from unittest.mock import MagicMock, patch
import DefFileGenerator.extractor as extractor_mod

class TestPDFPages(unittest.TestCase):
    @patch('DefFileGenerator.extractor.HAS_PDFPLUMBER', True)
    @patch('DefFileGenerator.extractor.pdfplumber')
    def test_extract_from_pdf_pages_validation(self, mock_pdfplumber):
        mock_pdf = MagicMock()
        mock_pdf.pages = [MagicMock(), MagicMock(), MagicMock()] # 3 pages
        for i, p in enumerate(mock_pdf.pages):
            p.page_number = i + 1
            p.extract_tables.return_value = [[["Address", "Name"], ["100", "Test"]]]

        mock_pdfplumber.open.return_value.__enter__.return_value = mock_pdf

        extractor = extractor_mod.Extractor()

        # Test valid single page
        list(extractor.extract_from_pdf("fake.pdf", pages=1))

        # Test out of range
        with self.assertLogs('root', level='WARNING') as cm:
            gen = extractor.extract_from_pdf("fake.pdf", pages=5)
            list(gen)
            self.assertTrue(any("Page 5 is out of range" in msg for msg in cm.output))

if __name__ == '__main__':
    unittest.main()
