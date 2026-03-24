import unittest
import os
import csv
from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.extractor = Extractor()

    def test_address_offset_simple(self):
        self.assertEqual(self.generator.apply_address_offset('30001', 10), '30011')
        self.assertEqual(self.generator.apply_address_offset('0x10', 10), '26')

    def test_address_offset_compound(self):
        self.assertEqual(self.generator.apply_address_offset('30001_20', 10), '30011_20')
        self.assertEqual(self.generator.apply_address_offset('30001_0_1', 10), '30011_0_1')

    def test_tag_sanitization(self):
        # Test the modular _process_name_and_tag which uses RE_TAG_CLEAN
        seen_names = {}
        seen_tags = {}
        tag = self.generator._process_name_and_tag("My @ Variable!", "", 1, seen_names, seen_tags)
        self.assertEqual(tag, "my_variable")

        # Test double underscore collapsing (implicit in RE_TAG_CLEAN.sub('_', ...))
        tag2 = self.generator._process_name_and_tag("My  Variable", "", 2, seen_names, seen_tags)
        self.assertEqual(tag2, "my_variable_1") # Incremented because it would be my_variable

    def test_tag_sanitization_advanced(self):
        seen_names = {}
        seen_tags = {}
        # Leading/trailing non-alphanumeric, and multiple in middle
        tag = self.generator._process_name_and_tag("!! My  @  Variable !!", "", 1, seen_names, seen_tags)
        self.assertEqual(tag, "my_variable")

    def test_compound_address_generation(self):
        table = [[
            {'Address': '30001', 'Length': '2', 'Name': 'Var1', 'Type': 'U32'}
        ]]
        mapped = self.extractor.map_and_clean(table)
        self.assertEqual(mapped[0]['Address'], '30001_2')

        table2 = [[
            {'Address': '30001', 'Length': '1', 'StartBit': '5', 'Name': 'Var2', 'Type': 'BITS'}
        ]]
        mapped2 = self.extractor.map_and_clean(table2)
        self.assertEqual(mapped2[0]['Address'], '30001_1_5')

if __name__ == '__main__':
    unittest.main()
