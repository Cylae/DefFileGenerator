import unittest
import os
import csv
import json
from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor

class TestNewFeatures(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.extractor = Extractor()

    def test_address_offset_logic(self):
        # Simple address
        self.assertEqual(self.generator.apply_address_offset("100", 10), "110")
        self.assertEqual(self.generator.apply_address_offset("0x64", 10), "110") # 0x64 is 100

        # Compound address (Address_Length)
        self.assertEqual(self.generator.apply_address_offset("100_20", 10), "110_20")

        # Compound address (Address_Length_StartBit)
        self.assertEqual(self.generator.apply_address_offset("100_1_5", 10), "110_1_5")

        # Negative result
        # Note: Generator.process_rows logs a warning but still produces the negative address
        self.assertEqual(self.generator.apply_address_offset("50", -100), "-50")

    def test_compound_address_generation(self):
        # Table with Address, Length, StartBit
        table = [
            {
                "Name": "Complex Var",
                "Address": "1000",
                "Length": "1",
                "StartBit": "8",
                "Type": "BITS"
            }
        ]
        # We need to mock the detection of len_col and bit_col in map_and_clean
        # In map_and_clean, it looks for columns containing 'length' or 'start bit'

        mapped = self.extractor.map_and_clean([table])
        self.assertEqual(mapped[0]["Address"], "1000_1_8")

        # Address_Length only
        table2 = [
            {
                "Name": "String Var",
                "Address": "2000",
                "Length": "10",
                "Type": "STRING"
            }
        ]
        mapped2 = self.extractor.map_and_clean([table2])
        self.assertEqual(mapped2[0]["Address"], "2000_10")

    def test_tag_sanitization(self):
        rows = [
            {"Name": "Voltage L1-L2!", "Address": "100", "Type": "U16"},
            {"Name": "Current @ Phase A", "Address": "101", "Type": "U16"},
            {"Name": "__Double Underscore__", "Address": "102", "Type": "U16"}
        ]
        processed = self.generator.process_rows(rows)

        # RE_TAG_CLEAN = re.compile(r'[^a-zA-Z0-9]+')
        # base_tag = RE_TAG_CLEAN.sub('_', name.lower().replace(' ', '_')).strip('_')

        # "Voltage L1-L2!" -> "voltage_l1_l2_" -> "voltage_l1_l2"
        self.assertEqual(processed[0]["Tag"], "voltage_l1_l2")

        # "Current @ Phase A" -> "current___phase_a" -> "current_phase_a"
        # Wait, RE_TAG_CLEAN replaces EACH non-alphanumeric with '_'.
        # re.sub(r'[^a-zA-Z0-9]+', '_', "current @ phase a".lower().replace(' ', '_'))
        # "current @ phase a" -> replace(' ', '_') -> "current_@_phase_a"
        # sub(r'[^a-zA-Z0-9]+', '_', "current_@_phase_a") -> "current_phase_a" (because _ is not in [a-zA-Z0-9])
        self.assertEqual(processed[1]["Tag"], "current_phase_a")

        # "__Double Underscore__" -> "double_underscore"
        self.assertEqual(processed[2]["Tag"], "double_underscore")

    def test_str_n_support(self):
        rows = [
            {"Name": "Serial", "Address": "3000", "Type": "STR20"}
        ]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]["Info3"], "STRING")
        self.assertEqual(processed[0]["Info2"], "3000_20")

if __name__ == "__main__":
    unittest.main()
