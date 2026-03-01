import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestRobustness(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_address_offset_dec(self):
        self.generator.address_offset = 1
        rows = [{'Name': 'Test', 'Address': '40001', 'Type': 'U16'}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '40000')

    def test_address_offset_hex(self):
        self.generator.address_offset = 1
        rows = [{'Name': 'Test', 'Address': '0x0002', 'Type': 'U16'}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '1')

    def test_negative_address_warning(self):
        self.generator.address_offset = 100
        rows = [{'Name': 'Test', 'Address': '50', 'Type': 'U16'}]
        with self.assertLogs(level='WARNING') as cm:
            processed = self.generator.process_rows(rows)
            self.assertEqual(processed[0]['Info2'], '-50')
            self.assertTrue(any("results in negative address -50" in output for output in cm.output))

    def test_composite_address_normalization(self):
        # STRING type: Address_Length
        rows = [{'Name': 'Str', 'Address': '0x10_10', 'Type': 'STRING'}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '16_10')

        # BITS type: Address_StartBit_NumBits
        rows = [{'Name': 'Bits', 'Address': '100_0_1', 'Type': 'BITS'}]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '100_0_1')

    def test_malformed_numeric_fallback(self):
        rows = [{
            'Name': 'Test',
            'Address': '100',
            'Type': 'U16',
            'Factor': 'invalid',
            'ScaleFactor': 'abc',
            'Offset': 'xyz'
        }]
        with self.assertLogs(level='WARNING') as cm:
            processed = self.generator.process_rows(rows)
            self.assertEqual(processed[0]['CoefA'], '1.000000') # Factor=1.0, ScaleFactor=0
            self.assertEqual(processed[0]['CoefB'], '0.000000')
            self.assertTrue(any("Invalid Factor 'invalid'" in output for output in cm.output))
            self.assertTrue(any("Invalid ScaleFactor 'abc'" in output for output in cm.output))
            self.assertTrue(any("Invalid Offset 'xyz'" in output for output in cm.output))

    def test_normalize_address_val_messy(self):
        self.assertEqual(self.generator.normalize_address_val("Address: 0x9C40"), "40000")
        self.assertEqual(self.generator.normalize_address_val("Reg 40001"), "40001")
        self.assertEqual(self.generator.normalize_address_val("10h"), "16")
        self.assertEqual(self.generator.normalize_address_val("40,001"), "40001")

if __name__ == '__main__':
    unittest.main()
