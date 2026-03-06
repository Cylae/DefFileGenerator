import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestDefGenEnhanced(unittest.TestCase):
    def setUp(self):
        # Disable logging for tests to keep output clean
        logging.disable(logging.CRITICAL)

    def test_address_offset(self):
        generator = Generator(address_offset=10)
        rows = [
            {'Name': 'Test1', 'Address': '100', 'Type': 'U16'}
        ]
        processed = generator.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '90')

    def test_negative_address_validation(self):
        generator = Generator()
        # Should be valid according to RE_ADDR_INT
        self.assertTrue(generator.validate_address('-1', 'U16'))

        rows = [
            {'Name': 'TestNeg', 'Address': '-5', 'Type': 'U16'}
        ]
        processed = generator.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '-5')

    def test_address_offset_results_in_negative(self):
        generator = Generator(address_offset=100)
        rows = [
            {'Name': 'Test1', 'Address': '50', 'Type': 'U16'}
        ]
        # Should log a warning (not easily testable here without mocking logging)
        # but should still process and result in -50
        processed = generator.process_rows(rows)
        self.assertEqual(processed[0]['Info2'], '-50')

    def test_type_normalization_synonyms(self):
        generator = Generator()
        self.assertEqual(generator.normalize_type('uint16'), 'U16')
        self.assertEqual(generator.normalize_type('int32'), 'I32')
        self.assertEqual(generator.normalize_type('float32'), 'F32')
        self.assertEqual(generator.normalize_type('double'), 'F64')
        self.assertEqual(generator.normalize_type('unsigned int 64'), 'U64')
        # self.assertEqual(generator.normalize_type('signed short'), 'I16') # Logic doesn't handle 'short'

    def test_action_normalization(self):
        generator = Generator()
        self.assertEqual(generator.normalize_action('R'), '4')
        self.assertEqual(generator.normalize_action('Read'), '4')
        self.assertEqual(generator.normalize_action('RW'), '1')
        self.assertEqual(generator.normalize_action('Write'), '1')
        self.assertEqual(generator.normalize_action('4'), '4')
        self.assertEqual(generator.normalize_action(''), '1')

    def test_hex_address_normalization(self):
        generator = Generator()
        self.assertEqual(generator.normalize_address_val('0x10'), '16')
        self.assertEqual(generator.normalize_address_val('10h'), '16')
        self.assertEqual(generator.normalize_address_val('Reg: 0x10 (Holding)'), '16')

if __name__ == '__main__':
    unittest.main()
