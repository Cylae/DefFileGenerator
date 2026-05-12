import unittest
from unittest.mock import MagicMock, patch
import logging
from DefFileGenerator.extractor import Extractor

class TestPdfValidation(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()

    def test_pdf_page_validation(self, mock_open=None):
        # We need to ensure the logger level allows WARNINGS
        with patch('pdfplumber.open') as mock_open:
            # Mock PDF with 5 pages
            mock_pdf = MagicMock()
            mock_pdf.pages = [MagicMock() for _ in range(5)]
            mock_open.return_value.__enter__.return_value = mock_pdf

            # Test out of range page
            with self.assertLogs(level='WARNING') as cm:
                list(self.extractor.extract_from_pdf('dummy.pdf', pages=[1, 6, 3]))

            self.assertTrue(any("Page 6 is out of range" in msg for msg in cm.output))

if __name__ == '__main__':
    unittest.main()
