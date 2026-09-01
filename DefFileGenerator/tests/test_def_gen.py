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

    def test_sanitize_csv_field_whitespace_prefixed_payloads(self):
        """Leading whitespace stripped by spreadsheet clients must not bypass escaping."""
        for payload in ("\t=1+1", "\r=1+1", "\n=1+1", " =HYPERLINK(\"x\")",
                        "\u00a0=1+1", "\x0b=1+1", "\x0c=1+1"):
            with self.subTest(payload=payload):
                self.assertEqual(self.generator.sanitize_csv_field(payload), "'" + payload)

    def test_sanitize_csv_field_fullwidth_and_pipe_triggers(self):
        """Full-width triggers and DDE pipe payloads are escaped."""
        for payload in ("\uff1d1+1", "\uff0bSUM", "\uff0d1", "\uff20SUM", "|dde"):
            with self.subTest(payload=payload):
                self.assertEqual(self.generator.sanitize_csv_field(payload), "'" + payload)

    def test_sanitize_csv_field_preserves_finite_numbers(self):
        """Genuine signed numbers stay verbatim; non-finite literals do not."""
        for value in ("-10.5", "+25", "-5", "1.5e3", "-1.5E-3", "40001", "+.5"):
            with self.subTest(value=value):
                self.assertEqual(self.generator.sanitize_csv_field(value), value)
        for value in ("-inf", "+nan", "-Infinity", "-1_000", " -10.5"):
            with self.subTest(value=value):
                self.assertEqual(self.generator.sanitize_csv_field(value), "'" + value)

    def test_action_code_10_constant_preserved(self):
        """FW 5.2.02 'Constant' variables use action code 10 and must survive."""
        rows = [{'Name': 'Rated Power', 'Tag': 'rated_power', 'RegisterType': '3',
                 'Address': '1', 'Type': 'U32', 'Action': '10'}]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '10')

    def test_unknown_action_still_falls_back(self):
        """An unrecognised action code keeps the register-class default."""
        rows = [
            {'Name': 'A', 'Tag': 'a', 'RegisterType': '3', 'Address': '10', 'Type': 'U16', 'Action': '99'},
            {'Name': 'B', 'Tag': 'b', 'RegisterType': '4', 'Address': '20', 'Type': 'U16', 'Action': 'bogus'},
        ]
        processed = list(self.generator.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')
        self.assertEqual(processed[1]['Action'], '4')

    def test_normalize_type_precedence_is_stable(self):
        """Specificity ordering of the synonym table must not regress."""
        cases = {
            'floatdouble': 'F64', 'double': 'F64', 'float': 'F32',
            'uint16': 'U16', 'signed integer 64': 'I64',
            'string 20': 'STR20', 'string': 'STRING',
            'uint32_wb': 'U32_WB', 'float 32 swap': 'F32_WB',
            'int8': 'I8', '': 'U16',
        }
        for raw, expected in cases.items():
            with self.subTest(raw=raw):
                self.assertEqual(self.generator.normalize_type(raw), expected)

    def test_normalize_address_leading_zero_contract(self):
        """Leading-zero tokens keep their legacy hex interpretation."""
        self.assertEqual(self.generator.normalize_address_val('0021'), '33')
        self.assertEqual(self.generator.normalize_address_val('012'), '18')
        self.assertEqual(self.generator.normalize_address_val('21'), '21')
        self.assertEqual(self.generator.normalize_address_val('-09'), '-09')

    def test_write_output_csv_is_atomic_on_failure(self):
        """A mid-stream error must not clobber an existing definition file."""
        import glob
        import os
        import tempfile
        fd, target = tempfile.mkstemp(suffix='.csv')
        os.close(fd)
        try:
            with open(target, 'w', encoding='utf-8') as f:
                f.write('EXISTING PRODUCTION DEFINITION\n')

            def exploding_rows():
                yield {'Info1': '3', 'Info2': '1', 'Info3': 'U16', 'Info4': '',
                       'Name': 'a', 'Tag': 'a', 'CoefA': '1.000000',
                       'CoefB': '0.000000', 'Unit': 'V', 'Action': '4'}
                raise RuntimeError('simulated extractor failure')

            with self.assertRaises(RuntimeError):
                Generator.write_output_csv(target, exploding_rows(), 'M', 'X')

            with open(target, encoding='utf-8') as f:
                self.assertEqual(f.read().strip(), 'EXISTING PRODUCTION DEFINITION')
            directory = os.path.dirname(target)
            self.assertEqual(glob.glob(os.path.join(directory, '.webdyn-*')), [])
        finally:
            os.unlink(target)

    def test_write_output_csv_replaces_on_success(self):
        """The happy path still fully replaces the target file."""
        import os
        import tempfile
        fd, target = tempfile.mkstemp(suffix='.csv')
        os.close(fd)
        try:
            rows = [{'Info1': '3', 'Info2': '40001', 'Info3': 'U16', 'Info4': '',
                     'Name': 'Active Power', 'Tag': 'active_power',
                     'CoefA': '1.000000', 'CoefB': '0.000000', 'Unit': 'W', 'Action': '4'}]
            Generator.write_output_csv(target, iter(rows), 'Huawei', 'SUN2000')
            with open(target, encoding='utf-8-sig') as f:
                content = f.read()
            self.assertIn('modbusRTU;Inverter;Huawei;SUN2000', content)
            self.assertIn('1;3;40001;U16;;Active Power;active_power', content)
        finally:
            os.unlink(target)

    def test_get_register_count_cache_consistency(self):
        """Memoised width lookup agrees with per-call STRING sizing."""
        self.assertEqual(self.generator.get_register_count('U16', '1'), 1)
        self.assertEqual(self.generator.get_register_count('U16', '2'), 1)
        self.assertEqual(self.generator.get_register_count('STRING', '30000_10'), 5)
        self.assertEqual(self.generator.get_register_count('STRING', '30000_20'), 10)
        self.assertEqual(self.generator.get_register_count('STRING', '30000'), 0)
        self.assertEqual(self.generator.get_register_count('U64', '1'), 4)

if __name__ == '__main__':
    unittest.main()
