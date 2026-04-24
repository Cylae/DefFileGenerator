import unittest
import logging
from DefFileGenerator.extractor import Extractor

class TestRobustness(unittest.TestCase):
    def setUp(self):
        logging.basicConfig(level=logging.DEBUG)

    def test_multipass_column_matching(self):
        # Pass 1: Normalized match (exact after stripping spaces/case)
        raw_data = [[{"Register Address": "100", "Variable Name": "TestVar", "Data Type": "U16"}]]
        extractor = Extractor()
        mapped = extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "100")
        self.assertEqual(mapped[0]["Name"], "TestVar")
        self.assertEqual(mapped[0]["Type"], "U16")

        # Pass 2: Substring match fallback
        # 'Current Scale' should match 'Factor' (contains 'scale')
        # 'Scale Factor' should match 'ScaleFactor' (contains 'scalefactor' in Pass 1)
        raw_data = [[{"Addr": "200", "Param": "Current", "Scale Factor": "2", "Current Scale": "0.1"}]]
        mapped = extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "200")
        self.assertEqual(mapped[0]["ScaleFactor"], "2")
        self.assertEqual(mapped[0]["Factor"], "0.1")

    def test_bits_handling_defaults(self):
        extractor = Extractor()

        # Simple address for BITS should default to _0_1
        raw_data = [[{"Address": "1000", "Name": "Status", "Type": "BITS"}]]
        mapped = extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "1000_0_1")

        # Explicit columns should override
        raw_data = [[{"Address": "2000", "Name": "Fault", "Type": "BITS", "StartBit": "5", "Length": "2"}]]
        mapped = extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "2000_5_2")

        # Compound address should be preserved
        raw_data = [[{"Address": "3000_8_1", "Name": "Warn", "Type": "BITS"}]]
        mapped = extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "3000_8_1")

    def test_address_with_thousands_separators(self):
        # Testing normalization of addresses with separators (e.g. 40,001)
        raw_data = [[{"Address": "40,001", "Name": "Power", "Type": "U16"}]]
        extractor = Extractor()
        mapped = extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "40001")

if __name__ == "__main__":
    unittest.main()
