import unittest
import time
import random
import logging
from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor

class TestExhaustive(unittest.TestCase):
    def setUp(self):
        self.generator = Generator()
        self.extractor = Extractor()
        # Suppress logging during massive tests to avoid I/O bottlenecks
        logging.getLogger().setLevel(logging.CRITICAL)

    def test_massive_dataset_performance(self):
        """Test performance and memory handling of process_rows with a huge dataset."""
        num_registers = 100000
        rows = []

        types = ['U16', 'I32', 'F32', 'U64', 'STRING', 'MAC', 'IPV6']
        actions = ['R', 'W', 'RW', '']

        for i in range(num_registers):
            reg_type = random.choice(types)
            address = str(30000 + i * 10)
            if reg_type == 'STRING':
                address = f"{address}_10"

            rows.append({
                'Name': f'Register_{i}',
                'Tag': f'reg_{i}',
                'RegisterType': 'Holding Register',
                'Address': address,
                'Type': reg_type,
                'Factor': '1',
                'Offset': '0',
                'Unit': 'V',
                'Action': random.choice(actions),
                'ScaleFactor': '0'
            })

        start_time = time.time()
        processed = self.generator.process_rows(rows)
        duration = time.time() - start_time

        self.assertEqual(len(processed), num_registers)
        print(f"\n[Performance] Processed {num_registers} registers in {duration:.2f} seconds.")
        # We expect it to be reasonably fast (e.g., < 5 seconds for 100k)
        self.assertLess(duration, 5.0, "Processing took too long!")

    def test_massive_overlaps(self):
        """Test the performance of the overlap detection algorithm with many overlaps."""
        num_registers = 50000
        rows = []

        # All using the exact same address to trigger massive overlap detection logic
        for i in range(num_registers):
            rows.append({
                'Name': f'Overlap_{i}',
                'Tag': f'ov_{i}',
                'RegisterType': '3',
                'Address': '40001',
                'Type': 'U32', # Takes 2 registers, 40001 and 40002
                'Factor': '1',
                'Offset': '0',
                'Unit': '',
                'Action': '4',
                'ScaleFactor': '0'
            })

        start_time = time.time()
        # Suppress warnings manually or capture them
        with self.assertLogs(level='WARNING') as log:
            processed = self.generator.process_rows(rows)
            duration = time.time() - start_time

        self.assertEqual(len(processed), num_registers)
        print(f"\n[Performance] Processed {num_registers} overlapping registers in {duration:.2f} seconds.")
        # Must be O(N) or close to it with dictionary check
        self.assertLess(duration, 15.0, "Overlap detection took too long!")
        self.assertGreater(len(log.output), 0, "No overlap warnings were logged!")

    def test_extreme_scalefactors(self):
        """Test with very large or very small ScaleFactors and Factors."""
        rows = [
            {'Name': 'Reg1', 'Tag': 't1', 'RegisterType': '3', 'Address': '100', 'Type': 'U16', 'Factor': '1e10', 'Offset': '0', 'ScaleFactor': '10', 'Action': ''},
            {'Name': 'Reg2', 'Tag': 't2', 'RegisterType': '3', 'Address': '101', 'Type': 'U16', 'Factor': '1e-10', 'Offset': '0', 'ScaleFactor': '-10', 'Action': ''},
            {'Name': 'Reg3', 'Tag': 't3', 'RegisterType': '3', 'Address': '102', 'Type': 'U16', 'Factor': '0.0000000001', 'Offset': '0', 'ScaleFactor': '5', 'Action': ''},
            {'Name': 'Reg4', 'Tag': 't4', 'RegisterType': '3', 'Address': '103', 'Type': 'U16', 'Factor': '123456789.987', 'Offset': '0', 'ScaleFactor': '-5', 'Action': ''},
        ]
        processed = self.generator.process_rows(rows)
        self.assertEqual(len(processed), 4)

    def test_invalid_data_types_and_addresses_recovery(self):
        """Test that the generator gracefully skips or handles completely broken rows."""
        rows = [
            {'Name': 'Broken1', 'Tag': 'b1', 'RegisterType': '3', 'Address': 'NOT_AN_ADDR', 'Type': 'U16', 'Factor': '', 'Offset': '', 'ScaleFactor': '', 'Action': ''},
            {'Name': 'Broken2', 'Tag': 'b2', 'RegisterType': '3', 'Address': '100', 'Type': 'NOT_A_TYPE', 'Factor': '', 'Offset': '', 'ScaleFactor': '', 'Action': ''},
            {'Name': '', 'Tag': 'b3', 'RegisterType': '3', 'Address': '101', 'Type': 'U16', 'Factor': '', 'Offset': '', 'ScaleFactor': '', 'Action': ''}, # No Name
        ]
        processed = self.generator.process_rows(rows)
        # Assuming missing Names or invalid types drop the row or leave it default
        # Actually, let's see how def_gen handles it. It might just skip or keep them.
        # We just want to make sure it doesn't crash.
        self.assertTrue(isinstance(processed, list))

    def test_massive_string_types(self):
        """Test massive STR<n> types up to STR1000 to check normalization and address expansion."""
        rows = []
        for i in range(1, 1001):
            rows.append({
                'Name': f'StrReg_{i}',
                'Tag': f'sreg_{i}',
                'RegisterType': '3',
                'Address': str(10000 + i * 1000),
                'Type': f'STR{i}',
                'Factor': '1', 'Offset': '0', 'Unit': '', 'Action': '4', 'ScaleFactor': '0'
            })

        start_time = time.time()
        processed = self.generator.process_rows(rows)
        duration = time.time() - start_time

        self.assertEqual(len(processed), 1000)
        self.assertEqual(processed[0]['Info3'], 'STRING')
        self.assertEqual(processed[0]['Info2'], '11000_1') # STR1 -> 1 char -> takes 1 register
        self.assertEqual(processed[-1]['Info3'], 'STRING')
        self.assertEqual(processed[-1]['Info2'], '1010000_1000')
        print(f"\n[Performance] Processed 1000 massive string types in {duration:.2f} seconds.")

if __name__ == '__main__':
    unittest.main()
