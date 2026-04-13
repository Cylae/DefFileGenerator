import unittest
import os
import csv
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator

class TestEnhancedFeatures(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()
        self.generator = Generator()
        self.xml_file = "test_registers.xml"

    def tearDown(self):
        if os.path.exists(self.xml_file):
            os.remove(self.xml_file)

    def test_compound_address_mapping(self):
        # Test Address_StartBit_Length construction
        raw_data = [[
            {"Addr": "100", "Start": "2", "Len": "1", "Name": "BitVar", "Type": "BITS"}
        ]]
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "100_2_1")

        # Test Address_Length for STRING
        raw_data = [[
            {"Addr": "200", "Len": "10", "Name": "StrVar", "Type": "STRING"}
        ]]
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "200_10")

    def test_xml_extraction_no_pandas(self):
        # Create dummy XML
        xml_content = """<registers>
            <row>
                <address>30001</address>
                <name>Voltage</name>
                <type>U16</type>
            </row>
            <row>
                <address>30002</address>
                <name>Current</name>
                <type>U16</type>
            </row>
        </registers>"""
        with open(self.xml_file, "w") as f:
            f.write(xml_content)

        data = self.extractor.extract_from_xml(self.xml_file)
        self.assertEqual(len(data), 1)
        self.assertEqual(len(data[0]), 2)
        self.assertEqual(data[0][0]["address"], "30001")
        self.assertEqual(data[0][1]["name"], "Current")

    def test_numeric_parsing(self):
        # Test comma as decimal separator
        self.assertEqual(self.generator._parse_numeric("0,1"), 0.1)

        # European style: 1.234,56
        self.assertEqual(self.generator._parse_numeric("1.234,56"), 1234.56)

        # US/UK style: 1,234.56
        self.assertEqual(self.generator._parse_numeric("1,234.56"), 1234.56)

        # Test fractions
        self.assertEqual(self.generator._parse_numeric("1/10"), 0.1)

    def test_address_offset_compound(self):
        addr = "100_2_1"
        offset = 1000
        result = self.generator.apply_address_offset(addr, offset)
        self.assertEqual(result, "1100_2_1")

if __name__ == "__main__":
    unittest.main()
