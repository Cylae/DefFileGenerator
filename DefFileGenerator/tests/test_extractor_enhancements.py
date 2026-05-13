import unittest
from unittest.mock import MagicMock, patch
from DefFileGenerator.extractor import Extractor
import logging

class TestExtractorEnhancements(unittest.TestCase):
    def test_pdf_page_range_validation(self):
        extractor = Extractor()
        # Mocking pdfplumber.open
        with patch('pdfplumber.open') as mock_open:
            mock_pdf = MagicMock()
            mock_pdf.pages = [MagicMock(), MagicMock()] # 2 pages
            mock_open.return_value.__enter__.return_value = mock_pdf

            with self.assertLogs(level='WARNING') as cm:
                list(extractor.extract_from_pdf("dummy.pdf", pages=[1, 3]))
                self.assertTrue(any("Page 3 is out of range (1-2). Skipping." in output for output in cm.output))

    def test_empty_generator_on_missing_deps(self):
        with patch('DefFileGenerator.extractor.HAS_OPENPYXL', False):
            extractor = Extractor()
            res = extractor.extract_from_excel("dummy.xlsx")
            self.assertEqual(list(res), [])

        with patch('DefFileGenerator.extractor.HAS_PDFPLUMBER', False):
            extractor = Extractor()
            res = extractor.extract_from_pdf("dummy.pdf")
            self.assertEqual(list(res), [])

        with patch('DefFileGenerator.extractor.HAS_DEFUSEDXML', False):
            extractor = Extractor()
            res = extractor.extract_from_xml("dummy.xml")
            self.assertEqual(list(res), [])

if __name__ == '__main__':
    unittest.main()
