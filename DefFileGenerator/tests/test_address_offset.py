import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestAddressOffset(unittest.TestCase):
    def test_address_offset(self):
        # Test basic offset subtraction
        gen = Generator(address_offset=10)
        rows = [
            {'Name': 'Test1', 'Address': '100', 'Type': 'U16', 'RegisterType': 'Holding Register'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '90')

        # Test hex address with offset
        rows = [
            {'Name': 'Test2', 'Address': '0x14', 'Type': 'U16', 'RegisterType': 'Holding Register'}
        ]
        # 0x14 is 20, offset 10 -> 10
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '10')

    def test_negative_address_warning(self):
        # Offset larger than address
        gen = Generator(address_offset=100)
        rows = [
            {'Name': 'Test3', 'Address': '10', 'Type': 'U16', 'RegisterType': 'Holding Register'}
        ]

        with self.assertLogs(level='WARNING') as cm:
            processed = gen.process_rows(rows)
            self.assertEqual(processed[0]['Info2'], '-90')
            self.assertTrue(any("results in negative address -90" in output for output in cm.output))

    def test_composite_address_offset(self):
        # Offset should only apply to the first part of Address_Length or Address_StartBit_NbBits
        gen = Generator(address_offset=100)

        # Address_Length
        rows = [
            {'Name': 'String1', 'Address': '1000_10', 'Type': 'STRING', 'RegisterType': 'Holding Register'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '900_10')

        # Address_StartBit_NbBits
        rows = [
            {'Name': 'Bits1', 'Address': '500_0_1', 'Type': 'BITS', 'RegisterType': 'Holding Register'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '400_0_1')

if __name__ == "__main__":
    unittest.main()
