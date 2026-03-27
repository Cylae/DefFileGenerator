import unittest
import logging
from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.extractor = Extractor()

    def test_address_offset_logic(self):
        # Simple address
        self.assertEqual(self.generator.apply_address_offset("100", 10), "110")
        # Hex address
        self.assertEqual(self.generator.apply_address_offset("0x64", 10), "110")
        # Compound address (STRING)
        self.assertEqual(self.generator.apply_address_offset("100_20", 10), "110_20")
        # Negative offset
        self.assertEqual(self.generator.apply_address_offset("100", -10), "90")

    def test_tag_sanitization(self):
        rows = [
            {'Name': 'Test @ Variable!', 'Address': '100', 'Type': 'U16'},
            {'Name': '   Space   Test   ', 'Address': '101', 'Type': 'U16'}
        ]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Tag'], 'test_variable')
        self.assertEqual(processed[1]['Tag'], 'space_test')

    def test_str_n_expansion(self):
        rows = [{'Name': 'StringTest', 'Address': '100', 'Type': 'STR20'}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Info3'], 'STRING')
        self.assertEqual(processed[0]['Info2'], '100_20')

    def test_overlap_detection_o1(self):
        rows = [
            {'Name': 'Var1', 'Address': '100', 'Type': 'U32', 'RegisterType': '3'}, # 100, 101
            {'Name': 'Var2', 'Address': '101', 'Type': 'U16', 'RegisterType': '3'}  # 101 (Overlap)
        ]
        with self.assertLogs(level='WARNING') as log:
            self.generator.process_rows(rows)
            self.assertTrue(any("Address overlap detected" in m for m in log.output))

    def test_offset_during_mapping(self):
        raw_tables = [[
            {'Name': 'Var1', 'Address': '100', 'Type': 'U16'}
        ]]
        mapped = self.extractor.map_and_clean(raw_tables, address_offset=50)
        self.assertEqual(mapped[0]['Address'], '150')

if __name__ == "__main__":
    unittest.main()
