import unittest
import os
import csv
import tempfile
import logging
from DefFileGenerator.def_gen import Generator

class TestValidation(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        # Disable logging for tests to keep output clean
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)

    def test_validate_address_range(self):
        # Valid range 0-65535
        self.assertTrue(self.generator.validate_address("0", "U16"))
        self.assertTrue(self.generator.validate_address("65535", "U16"))
        self.assertTrue(self.generator.validate_address("0x0", "U16"))
        self.assertTrue(self.generator.validate_address("0xFFFF", "U16"))

        # Invalid range
        self.assertFalse(self.generator.validate_address("-1", "U16"))
        self.assertFalse(self.generator.validate_address("65536", "U16"))
        self.assertFalse(self.generator.validate_address("0x10000", "U16"))

    def test_validate_address_str_synonyms(self):
        # STR20 is a synonym for STRING
        self.assertTrue(self.generator.validate_address("100_20", "STR20"))
        self.assertFalse(self.generator.validate_address("100", "STR20")) # Missing length part for STRING type

    def test_intelligent_action_defaulting(self):
        # Input (4) and Discrete (2) -> 4 (Read Only)
        rows_input = [{'Name': 'Input', 'Address': '100', 'RegisterType': 'input', 'Type': 'U16', 'Action': ''}]
        processed_input = list(self.generator.process_rows(rows_input))
        self.assertEqual(processed_input[0]['Action'], '4')

        rows_discrete = [{'Name': 'Discrete', 'Address': '101', 'RegisterType': 'discrete input', 'Type': 'U16', 'Action': ''}]
        processed_discrete = list(self.generator.process_rows(rows_discrete))
        self.assertEqual(processed_discrete[0]['Action'], '4')

        # Holding (3) and Coils (1) -> 1 (Read/Write)
        rows_holding = [{'Name': 'Holding', 'Address': '102', 'RegisterType': 'holding', 'Type': 'U16', 'Action': ''}]
        processed_holding = list(self.generator.process_rows(rows_holding))
        self.assertEqual(processed_holding[0]['Action'], '1')

        rows_coil = [{'Name': 'Coil', 'Address': '103', 'RegisterType': 'coil', 'Type': 'U16', 'Action': ''}]
        processed_coil = list(self.generator.process_rows(rows_coil))
        self.assertEqual(processed_coil[0]['Action'], '1')

    def test_validate_csv_format(self):
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.csv') as tmp:
            writer = csv.writer(tmp, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'Model', '', '', '', '', '', '', ''])
            # 1;3;100;U16;;Name;Tag;1.0;0.0;Unit;1
            writer.writerow(['1', '3', '100', 'U16', '', 'Var1', 'var1', '1.0', '0.0', 'V', '1'])
            tmp_path = tmp.name

        try:
            self.assertTrue(self.generator.validate_csv(tmp_path))
        finally:
            os.remove(tmp_path)

    def test_validate_csv_errors(self):
        # Test short row
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.csv') as tmp:
            writer = csv.writer(tmp, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'Model', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U16']) # Too short
            tmp_path = tmp.name
        try:
            self.assertFalse(self.generator.validate_csv(tmp_path))
        finally:
            os.remove(tmp_path)

        # Test overlap
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.csv') as tmp:
            writer = csv.writer(tmp, delimiter=';')
            writer.writerow(['modbusRTU', 'Inverter', 'MFG', 'Model', '', '', '', '', '', '', ''])
            writer.writerow(['1', '3', '100', 'U32', '', 'Var1', 'var1', '1.0', '0.0', 'V', '1'])
            writer.writerow(['2', '3', '101', 'U16', '', 'Var2', 'var2', '1.0', '0.0', 'V', '1'])
            tmp_path = tmp.name
        try:
            self.assertFalse(self.generator.validate_csv(tmp_path))
        finally:
            os.remove(tmp_path)

if __name__ == '__main__':
    unittest.main()
