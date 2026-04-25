import unittest
import os
import tempfile
import csv
from DefFileGenerator.extractor import Extractor

class TestRobustness(unittest.TestCase):
    def setUp(self):
        self.extractor = Extractor()

    def test_multipass_column_matching(self):
        # Test exact match priority over substring
        raw_data = [[
            {"Scale Factor": "10", "Factor": "1/10", "Name": "Var1", "Address": "100"}
        ]]
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["ScaleFactor"], "10")
        self.assertEqual(mapped[0]["Factor"], "0.1")

        # Test normalized match (ignoring spaces/underscores)
        raw_data = [[
            {"register_type": "3", "register__address": "200", "variable_name": "Var2"}
        ]]
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["RegisterType"], "3")
        self.assertEqual(mapped[0]["Address"], "200")
        self.assertEqual(mapped[0]["Name"], "Var2")

    def test_bits_address_defaulting(self):
        # Test BITS type defaults to Addr_0_1 if bit info missing
        raw_data = [[
            {"Address": "300", "Name": "BitsVar", "Type": "BITS"}
        ]]
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "300_0_1")

        # Test BITS type preserves existing compound address
        raw_data = [[
            {"Address": "400_5_2", "Name": "BitsVar2", "Type": "BITS"}
        ]]
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "400_5_2")

        # Test BITS type with separate StartBit and Length columns
        raw_data = [[
            {"Address": "500", "Name": "BitsVar3", "Type": "BITS", "StartBit": "8", "Length": "4"}
        ]]
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "500_8_4")

    def test_address_thousands_separator(self):
        # Test thousands separator normalization in addresses
        raw_data = [[
            {"Address": "40,001", "Name": "Var3", "Type": "U16"}
        ]]
        mapped = self.extractor.map_and_clean(raw_data)
        self.assertEqual(mapped[0]["Address"], "40001")

    def test_xml_xxe_protection(self):
        # Create a malicious XML file
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE root [
  <!ENTITY xxe "evil">
]>
<root>
  <row>
    <Name>Variable &xxe;</Name>
    <Address>1000</Address>
  </row>
</root>
"""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.xml', delete=False) as tf:
            tf.write(xml_content)
            temp_xml = tf.name

        try:
            # defusedxml should raise an EntitiesForbidden error when it encounters an entity
            from defusedxml import EntitiesForbidden
            with self.assertRaises(EntitiesForbidden):
                self.extractor.extract_from_xml(temp_xml)
        finally:
            if os.path.exists(temp_xml):
                os.remove(temp_xml)

if __name__ == "__main__":
    unittest.main()
