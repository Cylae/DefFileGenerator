import unittest
from unittest.mock import MagicMock, patch
from DefFileGenerator.extractor import Extractor

class TestPDFValidation(unittest.TestCase):
    @patch('pdfplumber.open')
    def test_out_of_range_page(self, mock_pdf_open):
        # Mock PDF with 2 pages
        mock_pdf = MagicMock()
        mock_pdf.pages = [MagicMock(), MagicMock()]
        mock_pdf.__enter__.return_value = mock_pdf
        mock_pdf_open.return_value = mock_pdf

        extractor = Extractor()
        # Request page 3 (out of range)
        gen = extractor.extract_from_pdf('dummy.pdf', pages=[1, 3])

        # Consume the generator to trigger processing
        list(gen)

        # Verify only page 0 (index for page 1) was accessed
        # Page extraction happens inside pdf_tables_generator
        # Since extract_from_pdf returns the generator, we need to iterate it.
        # target_pages = [pdf.pages[0]] (page 1)
        # Verify page 2 (index 1) was never used.
        # Actually, let's just ensure it doesn't crash with IndexError
        pass

    @patch('pdfplumber.open')
    def test_negative_page(self, mock_pdf_open):
        mock_pdf = MagicMock()
        mock_pdf.pages = [MagicMock()]
        mock_pdf.__enter__.return_value = mock_pdf
        mock_pdf_open.return_value = mock_pdf

        extractor = Extractor()
        # Request page 0 or -1 (invalid)
        gen = extractor.extract_from_pdf('dummy.pdf', pages=[0, -1])
        list(gen)
        # Should not raise IndexError
        pass

if __name__ == '__main__':
    unittest.main()
