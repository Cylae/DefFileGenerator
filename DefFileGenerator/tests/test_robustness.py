import unittest
import logging
import io
import os
import csv
from DefFileGenerator.def_gen import Generator, run_generator

class TestRobustness(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        # Suppress logging for tests unless needed
        logging.getLogger().setLevel(logging.ERROR)

    def test_address_offset_decimal(self):
        gen = Generator(address_offset=1)
        rows = [{'Name': 'Var1', 'Address': '101', 'Type': 'U16'}]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '100')

    def test_address_offset_hex(self):
        gen = Generator(address_offset=16)
        rows = [{'Name': 'Var1', 'Address': '0x20', 'Type': 'U16'}]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '16')

    def test_negative_address_warning(self):
        gen = Generator(address_offset=100)
        rows = [{'Name': 'Var1', 'Address': '50', 'Type': 'U16'}]
        with self.assertLogs(level='WARNING') as log:
            processed = gen.process_rows(rows)
            self.assertEqual(processed[0]['Info2'], '-50')
            self.assertTrue(any("results in negative address -50" in m for m in log.output))

    def test_composite_address_with_offset(self):
        gen = Generator(address_offset=1)
        # String Address_Length
        rows = [{'Name': 'Var1', 'Address': '101_10', 'Type': 'STRING'}]
        processed = gen.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '100_10')

    def test_robust_numeric_fallback(self):
        rows = [{
            'Name': 'Var1', 'Address': '100', 'Type': 'U16',
            'Factor': 'invalid', 'Offset': 'bad', 'ScaleFactor': 'none'
        }]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['CoefA'], '1.000000') # Default factor 1.0 * 10^0
        self.assertEqual(processed[0]['CoefB'], '0.000000') # Default offset 0.0

    def test_messy_address_normalization(self):
        # normalize_address_val uses regex to find first hex/dec word
        self.assertEqual(self.generator.normalize_address_val('Reg: 0x10 (Holding)'), '16')
        self.assertEqual(self.generator.normalize_address_val('40001 (Decimal)'), '40001')
        self.assertEqual(self.generator.normalize_address_val('A0h'), '160')

    def test_type_specificity(self):
        # float64 should match F64, not F32 via 'float'
        self.assertEqual(self.generator.normalize_type('float64'), 'F64')
        self.assertEqual(self.generator.normalize_type('float32'), 'F32')
        self.assertEqual(self.generator.normalize_type('float'), 'F32')

    def test_utf16_detection(self):
        # Create a dummy UTF-16 file with BOM
        test_file = 'test_utf16.csv'
        content = "Name,RegisterType,Address,Type\nVar1,Holding,100,U16"
        with open(test_file, 'w', encoding='utf-16') as f:
            f.write(content)

        output_file = 'test_utf16_out.csv'
        try:
            run_generator(test_file, output=output_file, manufacturer='M', model='M')
            self.assertTrue(os.path.exists(output_file))
            with open(output_file, 'r') as f:
                lines = f.readlines()
                self.assertTrue(len(lines) > 1)
                self.assertIn("Var1", lines[1])
        finally:
            if os.path.exists(test_file): os.remove(test_file)
            if os.path.exists(output_file): os.remove(output_file)

if __name__ == '__main__':
    unittest.main()
