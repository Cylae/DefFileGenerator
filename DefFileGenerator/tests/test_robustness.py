import unittest
from DefFileGenerator.extractor import Extractor

class TestRobustness(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()

    def test_multi_pass_matching(self):
        # Pass 1: Simple normalized match
        # 'reg_type' should match 'RegisterType' because:
        # src_norm = 'regtype'
        # target_norm = 'registertype'
        # Wait, 'regtype' is not 'registertype'.
        # But 'regtype' IS in COLUMN_MAPPING['RegisterType']

        # Let's try something that IS an exact normalized match
        raw_data = [[{"Register Type": "Holding", "addr": "100", "name": "Var1"}]]
        mapped = list(self.extractor.map_and_clean(raw_data))
        # 'Register Type' -> norm 'registertype' == target_norm 'registertype'
        self.assertEqual(mapped[0]["RegisterType"], "Holding")

    def test_bits_address_defaulting(self):
        raw_data = [[{"Name": "BitVar", "Address": "100", "Type": "BITS", "StartBit": "5", "Length": "1"}]]
        mapped = list(self.extractor.map_and_clean(raw_data))
        self.assertEqual(mapped[0]["Address"], "100_5_1")

    def test_string_address_defaulting(self):
        raw_data = [[{"Name": "StrVar", "Address": "200", "Type": "STR20", "Length": "20"}]]
        mapped = list(self.extractor.map_and_clean(raw_data))
        self.assertEqual(mapped[0]["Address"], "200_20")

    def test_address_normalization_no_corruption(self):
        # Hex address with underscore
        raw_data = [[{"Name": "HexVar", "Address": "0x10_2", "Type": "STRING"}]]
        mapped = list(self.extractor.map_and_clean(raw_data))
        self.assertEqual(mapped[0]["Address"], "16_2")

    def test_thousands_separator_normalization(self):
        raw_data = [[{"Name": "Var", "Address": "1,000", "Type": "U16"}]]
        mapped = list(self.extractor.map_and_clean(raw_data))
        self.assertEqual(mapped[0]["Address"], "1000")

    def test_empty_extraction(self):
        # Empty list of tables
        mapped = list(self.extractor.map_and_clean([]))
        self.assertEqual(len(mapped), 0)

        # List with empty table
        mapped = list(self.extractor.map_and_clean([[]]))
        self.assertEqual(len(mapped), 0)

        # Table with empty rows
        mapped = list(self.extractor.map_and_clean([[{}, {}]]))
        self.assertEqual(len(mapped), 0)

if __name__ == "__main__":
    unittest.main()
