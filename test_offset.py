from DefFileGenerator.def_gen import Generator
import unittest

class TestOffset(unittest.TestCase):
    def setUp(self):
        self.gen = Generator()

    def test_apply_offset(self):
        # Simple address
        self.assertEqual(self.gen.apply_address_offset("30001", 10), "30011")
        # Hex address
        self.assertEqual(self.gen.apply_address_offset("0x30", 16), "64")
        # Compound address (String)
        self.assertEqual(self.gen.apply_address_offset("30001_20", 10), "30011_20")
        # Compound address (Bits)
        self.assertEqual(self.gen.apply_address_offset("40001_0_1", 5), "40006_0_1")
        # Zero offset
        self.assertEqual(self.gen.apply_address_offset("100", 0), "100")
        # Negative offset
        self.assertEqual(self.gen.apply_address_offset("100", -10), "90")

if __name__ == '__main__':
    unittest.main()
