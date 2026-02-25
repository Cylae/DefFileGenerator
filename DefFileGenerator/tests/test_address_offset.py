import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestAddressOffset(unittest.TestCase):
    def test_offset_application(self):
        # Generator with offset 1
        gen = Generator(address_offset=1)
        rows = [
            {'Name': 'Var1', 'Address': '30001', 'Type': 'U16', 'RegisterType': '3'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '30000')

    def test_offset_composite_address(self):
        # Generator with offset 10
        gen = Generator(address_offset=10)
        rows = [
            {'Name': 'StrVar', 'Address': '30050_20', 'Type': 'STRING', 'RegisterType': '3'},
            {'Name': 'BitVar', 'Address': '30100_0_1', 'Type': 'BITS', 'RegisterType': '3'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '30040_20')
        self.assertEqual(processed[1]['Info2'], '30090_0_1')

    def test_negative_address_warning(self):
        # Generator with offset 100
        gen = Generator(address_offset=100)
        rows = [
            {'Name': 'VarLow', 'Address': '50', 'Type': 'U16', 'RegisterType': '3'}
        ]
        with self.assertLogs(level='WARNING') as log:
            processed = gen.process_rows(rows)
            self.assertEqual(processed[0]['Info2'], '-50')
            self.assertTrue(any("results in negative address -50" in m for m in log.output))

if __name__ == '__main__':
    unittest.main()
