import unittest
from DefFileGenerator.def_gen import Generator
import logging

class TestActionDefaults(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()

    def test_action_defaulting(self):
        test_cases = [
            # RegisterType, Expected Action
            ('Holding Register', '1'),
            ('Input Register', '4'),
            ('Coil', '1'),
            ('Discrete Input', '2'), # Discrete Input should default to 4 (Read Only) as per plan but let's see.
            # Wait, plan said '4' for Input Registers/Discrete Inputs.
            # Webdyn SunPM Info1: 1=Coil, 2=Discrete Input, 3=Holding, 4=Input Register.
            # Usually Coils/Holding are R/W (1), Discrete/Input are R only (4).
        ]

        for reg_type, expected_action in [
            ('Holding Register', '1'),
            ('Input Register', '4'),
            ('Coil', '1'),
            ('Discrete Input', '4'), # Based on plan: "4 for Input Registers/Discrete Inputs"
        ]:
            rows = [{
                'Name': f'Test_{reg_type}',
                'Address': '100',
                'RegisterType': reg_type,
                'Type': 'U16',
                'Action': ''
            }]
            processed = list(self.generator.process_rows(rows))
            with self.subTest(reg_type=reg_type):
                self.assertEqual(processed[0]['Action'], expected_action)

    def test_address_range_validation(self):
        with self.assertLogs(level='WARNING') as cm:
            list(self.generator.process_rows([{
                'Name': 'HighAddr',
                'Address': '70000',
                'Type': 'U16'
            }]))
            self.assertTrue(any("outside standard Modbus range" in output for output in cm.output))

if __name__ == '__main__':
    unittest.main()
