import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestActionDefaults(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.logger = logging.getLogger()
        self.original_level = self.logger.level

    def tearDown(self):
        self.logger.setLevel(self.original_level)

    def test_action_default_holding_register(self):
        rows = [{'name': 'test', 'address': '1', 'registertype': 'Holding Register', 'type': 'U16'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

    def test_action_default_coil(self):
        rows = [{'name': 'test', 'address': '1', 'registertype': 'Coil', 'type': 'U16'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

    def test_action_default_input_register(self):
        rows = [{'name': 'test', 'address': '1', 'registertype': 'Input Register', 'type': 'U16'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4')

    def test_action_default_discrete_input(self):
        rows = [{'name': 'test', 'address': '1', 'registertype': 'Discrete Input', 'type': 'U16'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4')

    def test_action_preserved_if_provided(self):
        rows = [{'name': 'test', 'address': '1', 'registertype': 'Holding Register', 'type': 'U16', 'action': '4'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4')

    def test_address_range_validation_warning(self):
        with self.assertLogs(level='WARNING') as cm:
            self.generator.validate_address("70000", "U16")
        self.assertTrue(any("outside standard Modbus range" in msg for msg in cm.output))

    def test_address_range_validation_ok(self):
        # Should not log warning
        self.logger.setLevel(logging.ERROR) # Only log ERRORS
        try:
            # This should NOT trigger a warning that we see if we set level high,
            # but assertLogs doesn't work that way easily.
            # We just call it and ensure it returns True.
            self.assertTrue(self.generator.validate_address("100", "U16"))
        finally:
            self.logger.setLevel(self.original_level)

if __name__ == '__main__':
    unittest.main()
