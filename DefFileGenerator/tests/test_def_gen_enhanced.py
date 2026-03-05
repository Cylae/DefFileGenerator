import unittest
import logging
import io
from DefFileGenerator.def_gen import Generator

class TestGeneratorEnhanced(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_normalize_type_synonyms(self):
        test_cases = [
            ('uint32', 'U32'),
            ('int16', 'I16'),
            ('float32', 'F32'),
            ('float64', 'F64'),
            ('double', 'F64'),
            ('uint16_w', 'U16_W'),
            ('int32_wb', 'I32_WB'),
            ('uint', 'U16'),
            ('boolean', 'BITS'),
            ('bool', 'BITS'),
            ('MAC', 'MAC'),
        ]
        for inp, expected in test_cases:
            with self.subTest(input=inp):
                self.assertEqual(self.generator.normalize_type(inp), expected)

    def test_address_offset(self):
        gen_with_offset = Generator(address_offset=1)
        rows = [{
            'Name': 'Test Var',
            'Address': '30001',
            'Type': 'U16',
        }]
        processed = gen_with_offset.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '30000')

    def test_negative_address_warning(self):
        gen_with_offset = Generator(address_offset=100)
        rows = [{
            'Name': 'Test Var',
            'Address': '10',
            'Type': 'U16',
        }]
        with self.assertLogs(level='WARNING') as log:
            processed = gen_with_offset.process_rows(rows)
            self.assertEqual(processed[0]['Info2'], '-90')
            self.assertTrue(any("results in negative address -90" in m for m in log.output))

    def test_normalize_address_val_hex_and_neg(self):
        self.assertEqual(self.generator.normalize_address_val('0x10'), '16')
        self.assertEqual(self.generator.normalize_address_val('-0x10'), '-16')
        self.assertEqual(self.generator.normalize_address_val('-10'), '-10')

if __name__ == '__main__':
    unittest.main()
