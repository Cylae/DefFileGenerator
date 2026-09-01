from unittest.mock import patch
from DefFileGenerator.def_gen import run_generator, GeneratorConfig
import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestGenerator(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        # Suppress logging during tests unless checking for logs
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)

    def test_intelligent_defaulting(self):
        rows = [
            {'Name': 'Holding', 'RegisterType': 'Holding', 'Address': '100', 'Type': 'U16'},
            {'Name': 'Input', 'RegisterType': 'Input', 'Address': '101', 'Type': 'U16'},
            {'Name': 'Coil', 'RegisterType': 'Coil', 'Address': '1', 'Type': 'U16'},
            {'Name': 'Discrete', 'RegisterType': 'Discrete', 'Address': '2', 'Type': 'U16'}
        ]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1') # Holding -> Read/Write
        self.assertEqual(processed[1]['Action'], '4') # Input -> Read Only
        self.assertEqual(processed[2]['Action'], '1') # Coil -> Read/Write
        self.assertEqual(processed[3]['Action'], '4') # Discrete -> Read Only

    def test_validate_address_range(self):
        self.assertTrue(self.generator.validate_address('0', 'U16'))
        self.assertTrue(self.generator.validate_address('65535', 'U16'))

        logging.disable(logging.NOTSET)
        with self.assertLogs(level='WARNING') as log:
            self.assertFalse(self.generator.validate_address('65536', 'U16'))
            self.assertTrue(any("out of standard Modbus range" in m for m in log.output))
        logging.disable(logging.CRITICAL)

    def test_normalize_address_val(self):
        self.assertEqual(self.generator.normalize_address_val('0x10'), '16')
        self.assertEqual(self.generator.normalize_address_val('10h'), '16')
        self.assertEqual(self.generator.normalize_address_val('10'), '10')
        self.assertEqual(self.generator.normalize_address_val('A0'), '160')
        self.assertEqual(self.generator.normalize_address_val('1,234'), '1234')

    def test_validate_address_invalid(self):
        self.assertFalse(self.generator.validate_address('30001_10', 'U16')) # U16 expects int
        self.assertFalse(self.generator.validate_address('xyz', 'U16')) # Not hex

    def test_get_register_count(self):
        self.assertEqual(self.generator.get_register_count('U16', '30000'), 1)
        self.assertEqual(self.generator.get_register_count('U32', '30000'), 2)
        self.assertEqual(self.generator.get_register_count('U64', '30000'), 4)
        self.assertEqual(self.generator.get_register_count('MAC', '30000'), 3)
        self.assertEqual(self.generator.get_register_count('IPV6', '30000'), 8)
        self.assertEqual(self.generator.get_register_count('STRING', '30000_10'), 5) # ceil(10/2)
        self.assertEqual(self.generator.get_register_count('STRING', '30000_11'), 6) # ceil(11/2)

    def test_process_rows_basic(self):
        rows = [{
            'Name': 'Test Var',
            'Tag': 'test_tag',
            'RegisterType': 'Holding Register',
            'Address': '30000',
            'Type': 'U16',
            'Factor': '1',
            'Offset': '0',
            'Unit': 'V',
            'Action': '4',
            'ScaleFactor': '0'
        }]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(len(processed), 1)
        self.assertEqual(processed[0]['Info1'], '3')
        self.assertEqual(processed[0]['Info3'], 'U16')
        self.assertEqual(processed[0]['CoefA'], '1.000000')

    def test_automatic_tag_generation(self):
        rows = [
            {'Name': 'Test Variable', 'Tag': '', 'RegisterType': '3', 'Address': '100', 'Type': 'U16'},
            {'Name': 'Test Variable', 'Tag': '', 'RegisterType': '3', 'Address': '101', 'Type': 'U16'}
        ]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Tag'], 'test_variable')
        self.assertEqual(processed[1]['Tag'], 'test_variable_1')

    def test_action_normalization(self):
        rows = [
            {'Name': 'Var1', 'Tag': 't1', 'RegisterType': '3', 'Address': '100', 'Type': 'U16', 'Action': 'R', 'Factor': '', 'Offset': '', 'Unit': '', 'ScaleFactor': ''},
            {'Name': 'Var2', 'Tag': 't2', 'RegisterType': '3', 'Address': '101', 'Type': 'U16', 'Action': 'RW', 'Factor': '', 'Offset': '', 'Unit': '', 'ScaleFactor': ''},
            {'Name': 'Var3', 'Tag': 't3', 'RegisterType': '3', 'Address': '102', 'Type': 'U16', 'Action': 'write', 'Factor': '', 'Offset': '', 'Unit': '', 'ScaleFactor': ''}
        ]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4') # R -> 4
        self.assertEqual(processed[1]['Action'], '1') # RW -> 1
        self.assertEqual(processed[2]['Action'], '1') # write -> 1

    def test_generate_template_modes(self):
        import os
        from DefFileGenerator.def_gen import generate_template
        # Test input mode
        out_input = "template_input.csv"
        generate_template(out_input, mode='input')
        with open(out_input, 'r') as f:
            content = f.read()
            self.assertIn("Name,Tag,RegisterType", content)
        os.remove(out_input)

        # Test definition mode
        out_def = "template_def.csv"
        generate_template(out_def, mode='definition')
        with open(out_def, 'r') as f:
            content = f.read()
            self.assertIn("modbusRTU;Inverter", content)
        os.remove(out_def)

    def test_sanitize_csv_field(self):
        self.assertEqual(self.generator.sanitize_csv_field(None), "")
        self.assertEqual(self.generator.sanitize_csv_field("Normal Text"), "Normal Text")
        self.assertEqual(self.generator.sanitize_csv_field("=1+1"), "'=1+1")
        self.assertEqual(self.generator.sanitize_csv_field("+cmd"), "'+cmd")
        self.assertEqual(self.generator.sanitize_csv_field("@SUM"), "'@SUM")
        self.assertEqual(self.generator.sanitize_csv_field("-10.5"), "-10.5")
        self.assertEqual(self.generator.sanitize_csv_field("+25"), "+25")
        self.assertEqual(self.generator.sanitize_csv_field("-text"), "'-text")

    def test_format_header_row(self):
        header = Generator._format_header_row('modbusRTU', 'Inverter', 'ACME', 'ModelX', 'forced')
        expected = ['modbusRTU', 'Inverter', 'ACME', 'ModelX', 'forced', '', '', '', '', '', '']
        self.assertEqual(header, expected)

    def test_format_data_row(self):
        row = {
            'Info1': '3',
            'Info2': '40001',
            'Info3': 'U16',
            'Info4': '',
            'Name': 'Voltage',
            'Tag': 'voltage',
            'CoefA': '1.000000',
            'CoefB': '0.000000',
            'Unit': 'V',
            'Action': '4'
        }
        data_row = Generator._format_data_row(1, row)
        expected = ['1', '3', '40001', 'U16', '', 'Voltage', 'voltage', '1.000000', '0.000000', 'V', '4']
        self.assertEqual(data_row, expected)

    def test_write_output_csv_io(self):
        import io
        buf = io.StringIO()
        rows = [{
            'Info1': '3',
            'Info2': '40001',
            'Info3': 'U16',
            'Info4': '',
            'Name': 'Active Power',
            'Tag': 'active_power',
            'CoefA': '1.000000',
            'CoefB': '0.000000',
            'Unit': 'W',
            'Action': '4'
        }]
        Generator.write_output_csv(buf, rows, 'ACME', 'ModelX')
        output = buf.getvalue()
        lines = output.strip().split('\n')
        self.assertEqual(lines[0], 'modbusRTU;Inverter;ACME;ModelX;;;;;;;')
        self.assertEqual(lines[1], '1;3;40001;U16;;Active Power;active_power;1.000000;0.000000;W;4')

if __name__ == '__main__':
    unittest.main()
