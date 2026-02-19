import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestAddressOffset(unittest.TestCase):
    def test_basic_offset(self):
        # 40001 with offset 1 should become 40000
        gen = Generator(address_offset=1)
        rows = [{'Name': 'Test', 'Address': '40001', 'Type': 'U16', 'RegisterType': 'Holding Register'}]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '40000')

    def test_hex_offset(self):
        # 0x10 (16) with offset 1 should become 15
        gen = Generator(address_offset=1)
        rows = [{'Name': 'Test', 'Address': '0x10', 'Type': 'U16', 'RegisterType': 'Holding Register'}]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '15')

    def test_negative_address_warning(self):
        # 10 with offset 20 should become -10 and log a warning
        gen = Generator(address_offset=20)
        rows = [{'Name': 'Test', 'Address': '10', 'Type': 'U16', 'RegisterType': 'Holding Register'}]

        with self.assertLogs(level='WARNING') as cm:
            processed = gen.process_rows(rows)
            self.assertEqual(processed[0]['Info2'], '-10')
            self.assertTrue(any("results in negative address -10" in output for output in cm.output))

    def test_composite_address_offset(self):
        # BITS: 30001_0_1 with offset 1 should become 30000_0_1
        gen = Generator(address_offset=1)
        rows = [{'Name': 'Test', 'Address': '30001_0_1', 'Type': 'BITS', 'RegisterType': 'Holding Register'}]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '30000_0_1')

        # STRING: 30010_20 with offset 10 should become 30000_20
        gen = Generator(address_offset=10)
        rows = [{'Name': 'Test', 'Address': '30010_20', 'Type': 'STRING', 'RegisterType': 'Holding Register'}]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '30000_20')

if __name__ == '__main__':
    unittest.main()
