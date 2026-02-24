import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestAddressOffset(unittest.TestCase):
    def test_address_offset_basic(self):
        # Subtract 1 from all addresses (convert 1-based to 0-based)
        gen = Generator(address_offset=1)
        rows = [
            {'Name': 'Var1', 'Address': '1', 'Type': 'U16'},
            {'Name': 'Var2', 'Address': '100', 'Type': 'U16'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '0')
        self.assertEqual(processed[1]['Info2'], '99')

    def test_address_offset_negative_warning(self):
        # Offset larger than address
        gen = Generator(address_offset=10)
        rows = [{'Name': 'Var1', 'Address': '5', 'Type': 'U16'}]
        with self.assertLogs(level='WARNING') as log:
            processed = gen.process_rows(rows)
            self.assertEqual(processed[0]['Info2'], '-5')
            self.assertTrue(any("results in negative address -5" in m for m in log.output))

    def test_address_offset_hex(self):
        gen = Generator(address_offset=1)
        rows = [{'Name': 'Var1', 'Address': '0x10', 'Type': 'U16'}] # 16 - 1 = 15
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '15')

    def test_address_offset_composite(self):
        gen = Generator(address_offset=40001)
        # 40001_10 -> (40001-40001)_10 -> 0_10
        rows = [{'Name': 'Var1', 'Address': '40001_10', 'Type': 'STRING'}]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '0_10')

if __name__ == '__main__':
    unittest.main()
