#!/usr/bin/env python3
import unittest
import logging
from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.extractor = Extractor()
        # Disable logging for tests
        logging.getLogger().setLevel(logging.ERROR)

    def test_apply_address_offset_simple(self):
        self.assertEqual(self.generator.apply_address_offset("30001", 10), "30011")
        self.assertEqual(self.generator.apply_address_offset("0x10", 1), "17") # 0x10 is 16

    def test_apply_address_offset_compound(self):
        # STRING: Addr_Len
        self.assertEqual(self.generator.apply_address_offset("30001_20", 10), "30011_20")
        # BITS: Addr_Len_Bit
        self.assertEqual(self.generator.apply_address_offset("30001_1_0", 10), "30011_1_0")
        # Hex base
        self.assertEqual(self.generator.apply_address_offset("0xA_10", 5), "15_10")

    def test_tag_sanitization(self):
        # Using _process_name_and_tag which uses RE_TAG_CLEAN
        seen_tags = {}
        tag = self.generator._process_name_and_tag(1, "Active Power (W)!!!", "", {}, seen_tags)
        self.assertEqual(tag, "active_power_w")

        # Test collapsing multiple underscores and stripping
        tag2 = self.generator._process_name_and_tag(2, "---Hello   World---", "", {}, seen_tags)
        self.assertEqual(tag2, "hello_world")

    def test_compound_address_extraction(self):
        # Simulating data with Address, Length, and StartBit columns
        table = [{
            'Name': 'Test Register',
            'Address': '30001',
            'Length': '2',
            'StartBit': '0',
            'Type': 'BITS'
        }]
        mapped = self.extractor.map_and_clean([table], address_offset=10)
        self.assertEqual(mapped[0]['Address'], '30011_2_0')

        # Just Address and Length
        table2 = [{
            'Name': 'Test String',
            'Address': '30050',
            'Length': '20',
            'Type': 'STRING'
        }]
        mapped2 = self.extractor.map_and_clean([table2], address_offset=0)
        self.assertEqual(mapped2[0]['Address'], '30050_20')

    def test_address_overlap_bits(self):
        # BITS at the same address should NOT trigger a warning (in my implementation it should be allowed)
        # We check this by seeing if the warning is logged (we can't easily check logging here without more setup,
        # but we can check if it finishes without error)
        rows = [
            {'Name': 'Bit 0', 'Address': '30001_1_0', 'Type': 'BITS', 'RegisterType': 'Holding'},
            {'Name': 'Bit 1', 'Address': '30001_1_1', 'Type': 'BITS', 'RegisterType': 'Holding'}
        ]
        processed = self.generator.process_rows(rows)
        self.assertEqual(len(processed), 2)

    def test_str_n_convenience_type(self):
        rows = [{'Name': 'String Var', 'Address': '30001', 'Type': 'STR20'}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Info3'], 'STRING')
        self.assertEqual(processed[0]['Info2'], '30001_20')

if __name__ == '__main__':
    unittest.main()
