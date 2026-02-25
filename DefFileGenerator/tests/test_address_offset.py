import unittest
import logging
from io import StringIO
from DefFileGenerator.def_gen import Generator

class TestAddressOffset(unittest.TestCase):
    def setUp(self):
        self.log_output = StringIO()
        self.handler = logging.StreamHandler(self.log_output)
        logging.getLogger().addHandler(self.handler)
        logging.getLogger().setLevel(logging.WARNING)

    def tearDown(self):
        logging.getLogger().removeHandler(self.handler)

    def test_simple_offset(self):
        gen = Generator(address_offset=10)
        rows = [
            {'Name': 'Var1', 'Address': '100', 'Type': 'U16'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '90')

    def test_composite_offset_string(self):
        gen = Generator(address_offset=10)
        rows = [
            {'Name': 'StrVar', 'Address': '100_20', 'Type': 'STRING'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '90_20')

    def test_composite_offset_bits(self):
        gen = Generator(address_offset=10)
        rows = [
            {'Name': 'BitVar', 'Address': '100_0_1', 'Type': 'BITS'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '90_0_1')

    def test_negative_address_warning(self):
        gen = Generator(address_offset=200)
        rows = [
            {'Name': 'NegVar', 'Address': '100', 'Type': 'U16'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '-100')

        log_content = self.log_output.getvalue()
        self.assertIn("Line 2: Address 100 with offset 200 results in negative address -100", log_content)

    def test_zero_offset(self):
        gen = Generator(address_offset=0)
        rows = [
            {'Name': 'Var1', 'Address': '100', 'Type': 'U16'}
        ]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '100')

if __name__ == '__main__':
    unittest.main()
