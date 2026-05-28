import unittest
import logging
import io
import os
import csv
from DefFileGenerator.def_gen import Generator, peek_generator
from DefFileGenerator.extractor import Extractor

class TestNewFeatures(unittest.TestCase):
    def test_peek_generator(self):
        def my_gen():
            yield 1
            yield 2

        ok, it = peek_generator(my_gen())
        self.assertTrue(ok)
        self.assertEqual(list(it), [1, 2])

        ok, it = peek_generator([])
        self.assertFalse(ok)
        self.assertEqual(list(it), [])

        ok, it = peek_generator(None)
        self.assertFalse(ok)
        self.assertEqual(list(it), [])

    def test_intelligent_action_defaulting(self):
        generator = Generator()
        rows = [
            {'Name': 'InputReg', 'RegisterType': 'Input Register', 'Address': '100', 'Type': 'U16'},
            {'Name': 'DiscInput', 'RegisterType': 'Discrete Input', 'Address': '101', 'Type': 'U16'},
            {'Name': 'HoldReg', 'RegisterType': 'Holding Register', 'Address': '102', 'Type': 'U16'},
            {'Name': 'Coil', 'RegisterType': 'Coil', 'Address': '103', 'Type': 'U16'},
        ]
        processed = list(generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4') # Input Register -> Read Only
        self.assertEqual(processed[1]['Action'], '4') # Discrete Input -> Read Only
        self.assertEqual(processed[2]['Action'], '1') # Holding Register -> Read/Write
        self.assertEqual(processed[3]['Action'], '1') # Coil -> Read/Write

    def test_pdf_pages_normalization(self):
        from unittest.mock import patch, MagicMock
        extractor = Extractor()

        with patch('pdfplumber.open') as mock_open:
            mock_pdf = MagicMock()
            mock_pdf.pages = [MagicMock(), MagicMock(), MagicMock()]
            mock_open.return_value.__enter__.return_value = mock_pdf

            # Test string input
            # Set up mock tables so the generator yields something
            mock_pdf.pages[0].extract_tables.return_value = [['Header', 'Data'], ['Name', 'Value']]
            mock_pdf.pages[2].extract_tables.return_value = [['Header', 'Data'], ['Name', 'Value']]

            # Consume the outer generator
            tables_gen = extractor.extract_from_pdf("dummy.pdf", pages="1,3")
            list(tables_gen)

            self.assertEqual(mock_pdf.pages[0].extract_tables.call_count, 1)
            self.assertEqual(mock_pdf.pages[1].extract_tables.call_count, 0)
            self.assertEqual(mock_pdf.pages[2].extract_tables.call_count, 1)

            # Test out of range
            for p in mock_pdf.pages: p.extract_tables.reset_mock()
            with self.assertLogs(level='WARNING') as log:
                list(extractor.extract_from_pdf("dummy.pdf", pages="1,5"))
                self.assertTrue(any("Page 5 is out of range" in m for m in log.output))

    def test_address_range_validation_log(self):
        generator = Generator()
        # Test negative address after offset
        with self.assertLogs(level='WARNING') as log:
            generator.apply_address_offset('10', -20, name='TestVar')
            self.assertTrue(any("Address offset -20 results in negative address -10 for 'TestVar'" in m for m in log.output))

if __name__ == '__main__':
    unittest.main()
