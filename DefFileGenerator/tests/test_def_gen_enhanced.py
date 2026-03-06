import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestGeneratorEnhanced(unittest.TestCase):
    def setUp(self):
        # Set up a generator with an address offset
        self.generator = Generator(address_offset=1)

    def test_address_offset(self):
        rows = [{
            'Name': 'Test Var',
            'Address': '30001',
            'Type': 'U16'
        }]
        processed = self.generator.process_rows(rows)
        # 30001 - 1 = 30000
        self.assertEqual(processed[0]['Info2'], '30000')

    def test_negative_address_warning(self):
        rows = [{
            'Name': 'Negative Var',
            'Address': '0',
            'Type': 'U16'
        }]
        with self.assertLogs(level='WARNING') as log:
            processed = self.generator.process_rows(rows)
            # 0 - 1 = -1
            self.assertEqual(processed[0]['Info2'], '-1')
            self.assertTrue(any("Resulting address -1 is negative" in m for m in log.output))

    def test_complex_type_normalization(self):
        gen = Generator()
        # Test specificity (float64 before float)
        self.assertEqual(gen.normalize_type("float64"), "F64")
        self.assertEqual(gen.normalize_type("float"), "F32")

        # Test synonyms
        self.assertEqual(gen.normalize_type("unsigned int 64"), "U64")
        self.assertEqual(gen.normalize_type("signed integer 32"), "I32")

        # Test preservation of suffixes
        self.assertEqual(gen.normalize_type("uint16_W"), "U16_W")
        self.assertEqual(gen.normalize_type("int32_WB"), "I32_WB")

    def test_robust_address_normalization(self):
        gen = Generator()
        # Hex with h suffix
        self.assertEqual(gen.normalize_address_val("10h"), "16")
        # Hex with 0x prefix
        self.assertEqual(gen.normalize_address_val("0x10"), "16")
        # Embedded address
        self.assertEqual(gen.normalize_address_val("Reg: 0x10"), "16")
        # Thousands separator
        self.assertEqual(gen.normalize_address_val("40,001"), "40001")
        # Negative address
        self.assertEqual(gen.normalize_address_val("-10"), "-10")

    def test_action_normalization_synonyms(self):
        gen = Generator()
        self.assertEqual(gen.normalize_action("Read"), "4")
        self.assertEqual(gen.normalize_action("Write"), "1")
        self.assertEqual(gen.normalize_action("RW"), "1")
        self.assertEqual(gen.normalize_action("R"), "4")

if __name__ == '__main__':
    unittest.main()
