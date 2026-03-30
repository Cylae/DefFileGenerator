import unittest
import logging
from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.extractor = Extractor()
        # Suppress logging for cleaner tests
        logging.getLogger().setLevel(logging.ERROR)

    def test_apply_address_offset_decimal(self):
        self.assertEqual(self.generator.apply_address_offset("40001", 10), "40011")
        self.assertEqual(self.generator.apply_address_offset("40001", 0), "40001")
        self.assertEqual(self.generator.apply_address_offset("40001", -10), "39991")

    def test_apply_address_offset_hex(self):
        # 0x0001 + 10 = 11
        self.assertEqual(self.generator.apply_address_offset("0x0001", 10), "11")
        # 0x10 + 10 = 16 + 10 = 26
        self.assertEqual(self.generator.apply_address_offset("0x10", 10), "26")
        # 10h + 10 = 16 + 10 = 26
        self.assertEqual(self.generator.apply_address_offset("10h", 10), "26")

    def test_apply_address_offset_compound(self):
        # String: Addr_Len
        self.assertEqual(self.generator.apply_address_offset("30001_20", 10), "30011_20")
        # Bits: Addr_StartBit_Len
        self.assertEqual(self.generator.apply_address_offset("30001_0_16", 10), "30011_0_16")
        # Hex compound
        self.assertEqual(self.generator.apply_address_offset("0x0001_20", 10), "11_20")

    def test_tag_sanitization(self):
        # Test tag generation in process_rows
        rows = [
            {"Name": "Test Name !@#", "Address": "40001", "Type": "U16"}
        ]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]["Tag"], "test_name")

        # Test collapse multiple underscores
        rows = [
            {"Name": "Test   Multiple   Spaces", "Address": "40001", "Type": "U16"}
        ]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]["Tag"], "test_multiple_spaces")

        # Test strip leading/trailing underscores
        rows = [
            {"Name": "  _Test_  ", "Address": "40001", "Type": "U16"}
        ]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]["Tag"], "test")

    def test_str_n_handling(self):
        rows = [
            {"Name": "String Test", "Address": "40001", "Type": "STR20"}
        ]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]["Info3"], "STRING")
        self.assertEqual(processed[0]["Info2"], "40001_20")

    def test_offset_during_mapping(self):
        # Test offset application in Extractor.map_and_clean
        tables = [[
            {"Name": "Voltage", "Address": "40001", "Type": "U16"}
        ]]
        mapped = self.extractor.map_and_clean(tables, address_offset=50)
        self.assertEqual(mapped[0]["Address"], "40051")

    def test_address_overlap_o1(self):
        # This is more of a logic check than a performance test
        rows = [
            {"Name": "Var1", "Address": "40001", "Type": "U32"}, # Uses 40001, 40002
            {"Name": "Var2", "Address": "40002", "Type": "U16"}  # Overlaps 40002
        ]
        with self.assertLogs(level='WARNING') as cm:
            self.generator.process_rows(rows)
            self.assertTrue(any("overlap detected" in output for output in cm.output))

if __name__ == "__main__":
    unittest.main()
