import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestAddressOffset(unittest.TestCase):
    def setUp(self):
        # We will initialize generator with different offsets in tests
        pass

    def test_positive_offset(self):
        gen = Generator(address_offset=1)
        rows = [
            {'Name': 'Var1', 'Address': '30001', 'Type': 'U16'}
        ]
        processed = gen.process_rows(rows)
        # 30001 - 1 = 30000
        self.assertEqual(processed[0]['Info2'], '30000')

    def test_hex_with_offset(self):
        gen = Generator(address_offset=16)
        rows = [
            {'Name': 'Var1', 'Address': '0x11', 'Type': 'U16'}
        ]
        processed = gen.process_rows(rows)
        # 0x11 (17) - 16 = 1
        self.assertEqual(processed[0]['Info2'], '1')

    def test_negative_result_warning(self):
        gen = Generator(address_offset=100)
        rows = [
            {'Name': 'Var1', 'Address': '50', 'Type': 'U16'}
        ]
        with self.assertLogs(level='WARNING') as log:
            processed = gen.process_rows(rows)
            # 50 - 100 = -50
            self.assertEqual(processed[0]['Info2'], '-50')
            self.assertTrue(any("results in negative address -50" in m for m in log.output))

    def test_composite_address_offset(self):
        gen = Generator(address_offset=1)
        rows = [
            {'Name': 'StrVar', 'Address': '30001_10', 'Type': 'STRING'},
            {'Name': 'BitVar', 'Address': '30001_0_1', 'Type': 'BITS'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '30000_10')
        self.assertEqual(processed[1]['Info2'], '30000_0_1')

if __name__ == '__main__':
    unittest.main()
