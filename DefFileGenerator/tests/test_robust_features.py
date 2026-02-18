import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestRobustFeatures(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_address_offset(self):
        gen = Generator(address_offset=1)
        rows = [{
            'Name': 'Var1',
            'Address': '30001',
            'Type': 'U16',
            'RegisterType': 'Holding Register'
        }]
        processed = gen.process_rows(rows)
        # 30001 - 1 = 30000
        self.assertEqual(processed[0]['Info2'], '30000')

    def test_address_offset_hex(self):
        gen = Generator(address_offset=16)
        rows = [{
            'Name': 'Var1',
            'Address': '0x11', # 17
            'Type': 'U16',
            'RegisterType': 'Holding Register'
        }]
        processed = gen.process_rows(rows)
        # 17 - 16 = 1
        self.assertEqual(processed[0]['Info2'], '1')

    def test_negative_address_warning(self):
        gen = Generator(address_offset=100)
        rows = [{
            'Name': 'Var1',
            'Address': '50',
            'Type': 'U16',
            'RegisterType': 'Holding Register'
        }]
        with self.assertLogs(level='WARNING') as log:
            processed = gen.process_rows(rows)
            self.assertEqual(processed[0]['Info2'], '-50')
            self.assertTrue(any("results in negative address -50" in m for m in log.output))

    def test_fractional_factor(self):
        rows = [{
            'Name': 'Var1',
            'Address': '100',
            'Type': 'U16',
            'Factor': '1/10',
            'ScaleFactor': '0'
        }]
        processed = self.generator.process_rows(rows)
        self.assertEqual(processed[0]['CoefA'], '0.100000')

    def test_normalize_type_synonyms(self):
        self.assertEqual(self.generator.normalize_type('uint16'), 'U16')
        self.assertEqual(self.generator.normalize_type('float32'), 'F32')
        self.assertEqual(self.generator.normalize_type('Double'), 'F64')
        self.assertEqual(self.generator.normalize_type('STR20'), 'STR20')
        self.assertEqual(self.generator.normalize_type('uint32(be)'), 'U32')

    def test_normalize_address_val_negative(self):
        self.assertEqual(self.generator.normalize_address_val('-10'), '-10')
        self.assertEqual(self.generator.normalize_address_val('-0x10'), '-16')

if __name__ == '__main__':
    unittest.main()
