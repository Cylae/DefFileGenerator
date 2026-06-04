import unittest
import logging
from DefFileGenerator.def_gen import Generator

class TestActionSynonyms(unittest.TestCase):
    def setUp(self):
        self.gen = Generator()
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)

    def test_action_defaulting(self):
        # Input Register (4) -> Read Only (4)
        rows = [{'Name': 'Test', 'Address': '1', 'RegisterType': 'Input Register', 'Type': 'U16'}]
        processed = list(self.gen.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '4')

        # Holding Register (3) -> Read/Write (1)
        rows = [{'Name': 'Test', 'Address': '1', 'RegisterType': 'Holding Register', 'Type': 'U16'}]
        processed = list(self.gen.process_rows(rows))
        self.assertEqual(processed[0]['Action'], '1')

    def test_action_synonyms(self):
        synonyms_ro = ['R', 'READ', 'RO', 'READ-ONLY', 'READ ONLY']
        for syn in synonyms_ro:
            rows = [{'Name': 'Test', 'Address': '1', 'Action': syn}]
            processed = list(self.gen.process_rows(rows))
            self.assertEqual(processed[0]['Action'], '4', f"Failed for synonym: {syn}")

        synonyms_rw = ['RW', 'W', 'WRITE', 'READ/WRITE', 'READ-WRITE', 'R/W']
        for syn in synonyms_rw:
            rows = [{'Name': 'Test', 'Address': '1', 'Action': syn}]
            processed = list(self.gen.process_rows(rows))
            self.assertEqual(processed[0]['Action'], '1', f"Failed for synonym: {syn}")

if __name__ == '__main__':
    unittest.main()
