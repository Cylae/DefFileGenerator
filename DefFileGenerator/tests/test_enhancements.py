import unittest
import logging
from DefFileGenerator.def_gen import Generator
from io import StringIO

class TestDefGenEnhancements(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_address_range_validation(self):
        with self.assertLogs(level='WARNING') as cm:
            self.generator.validate_address("65536", "U16")
            self.assertTrue(any("Address 65536 is outside standard Modbus range" in output for output in cm.output))

        with self.assertLogs(level='WARNING') as cm:
            self.generator.validate_address("-1", "U16")
            self.assertTrue(any("Address -1 is outside standard Modbus range" in output for output in cm.output))

    def test_intelligent_action_defaulting(self):
        # Input Register (4) should default to '4'
        rows = [{'name': 'Reg1', 'address': '100', 'type': 'U16', 'registertype': 'input register'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4')

        # Holding Register (3) should default to '1'
        rows = [{'name': 'Reg2', 'address': '200', 'type': 'U16', 'registertype': 'holding register'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

    def test_summary_logging(self):
        processed_rows = [
            {'Info1': '3', 'Info2': '100', 'Info3': 'U16', 'Info4': '', 'Name': 'N1', 'Tag': 'T1', 'CoefA': '1', 'CoefB': '0', 'Unit': 'V', 'Action': '1'},
            {'Info1': '4', 'Info2': '200', 'Info3': 'U16', 'Info4': '', 'Name': 'N2', 'Tag': 'T2', 'CoefA': '1', 'CoefB': '0', 'Unit': 'A', 'Action': '4'}
        ]
        output = StringIO()
        with self.assertLogs(level='INFO') as cm:
            self.generator.write_output_csv(output, processed_rows, "Mfg", "Model")
            self.assertTrue(any("Processed registers: Holding Registers: 1, Input Registers: 1" in output for output in cm.output))

if __name__ == '__main__':
    unittest.main()
