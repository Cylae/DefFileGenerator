import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestRobustness(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_address_offset_decimal(self):
        gen = Generator(address_offset=1)
        rows = [{'Name': 'Test', 'Address': '101', 'Type': 'U16'}]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '100')

    def test_address_offset_hex(self):
        gen = Generator(address_offset=1)
        rows = [{'Name': 'Test', 'Address': '0x0065', 'Type': 'U16'}] # 101 - 1 = 100
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '100')

    def test_address_offset_composite(self):
        gen = Generator(address_offset=1)
        # String: Address_Length
        rows = [{'Name': 'Test', 'Address': '101_20', 'Type': 'STRING'}]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '100_20')

    def test_negative_address_warning(self):
        gen = Generator(address_offset=100)
        rows = [{'Name': 'Test', 'Address': '50', 'Type': 'U16'}]
        with self.assertLogs(level='WARNING') as log:
            processed = gen.process_rows(rows)
            self.assertEqual(processed[0]['Info2'], '-50')
            self.assertTrue(any("results in negative address -50" in m for m in log.output))

    def test_malformed_numeric_fields(self):
        rows = [{
            'Name': 'Test', 'Address': '100', 'Type': 'U16',
            'Factor': 'abc', 'Offset': 'xyz', 'ScaleFactor': 'foo'
        }]
        with self.assertLogs(level='WARNING') as log:
            processed = self.generator.process_rows(rows)
            # Should use defaults: Factor=1.0, Offset=0.0, ScaleFactor=0
            self.assertEqual(processed[0]['CoefA'], '1.000000')
            self.assertEqual(processed[0]['CoefB'], '0.000000')
            self.assertTrue(any("Invalid Factor" in m for m in log.output))
            self.assertTrue(any("Invalid Offset" in m for m in log.output))
            self.assertTrue(any("Invalid ScaleFactor" in m for m in log.output))

    def test_normalize_type_robustness(self):
        self.assertEqual(self.generator.normalize_type("Unsigned Int 16"), "U16")
        self.assertEqual(self.generator.normalize_type("signed_int_32"), "I32")
        self.assertEqual(self.generator.normalize_type("Float 64"), "F64")
        self.assertEqual(self.generator.normalize_type("bool"), "BITS")

if __name__ == '__main__':
    unittest.main()
