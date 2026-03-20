import unittest
from DefFileGenerator.def_gen import Generator

class TestAddressOffset(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_apply_address_offset_simple(self):
        self.assertEqual(self.generator.apply_address_offset('30001', 10), '30011')
        self.assertEqual(self.generator.apply_address_offset('30001', -1), '30000')

    def test_apply_address_offset_hex(self):
        self.assertEqual(self.generator.apply_address_offset('0x10', 10), '26')
        self.assertEqual(self.generator.apply_address_offset('10h', 10), '26')

    def test_apply_address_offset_compound(self):
        self.assertEqual(self.generator.apply_address_offset('30001_20', 10), '30011_20')
        self.assertEqual(self.generator.apply_address_offset('30001_0_1', 10), '30011_0_1')

    def test_apply_address_offset_zero(self):
        self.assertEqual(self.generator.apply_address_offset('30001', 0), '30001')

if __name__ == '__main__':
    unittest.main()
