import unittest
import itertools
import logging
from io import StringIO
from DefFileGenerator.extractor import Extractor

class TestPeekingLogic(unittest.TestCase):
    def test_peeking_empty_generator(self):
        def empty_gen():
            if False:
                yield {}

        gen = empty_gen()
        with self.assertRaises(StopIteration):
            first = next(gen)

    def test_peeking_non_empty_generator(self):
        def non_empty_gen():
            yield {"a": 1}
            yield {"b": 2}

        gen = non_empty_gen()
        first = next(gen)
        self.assertEqual(first, {"a": 1})

        # Chain it back
        chained = itertools.chain([first], gen)
        items = list(chained)
        self.assertEqual(len(items), 2)
        self.assertEqual(items[0], {"a": 1})
        self.assertEqual(items[1], {"b": 2})

    def test_peeking_logic_usage_pattern(self):
        # This simulates the pattern used in doc_to_webdyn.py and main.py
        def my_gen():
            yield {"RegisterType": "Holding", "Address": "100", "Name": "Test"}
            yield {"RegisterType": "Holding", "Address": "101", "Name": "Test2"}

        gen = my_gen()
        try:
            first = next(gen)
            mapped = itertools.chain([first], gen)
        except StopIteration:
            self.fail("Generator should not be empty")

        results = list(mapped)
        self.assertEqual(len(results), 2)
        self.assertEqual(results[0]["Name"], "Test")
        self.assertEqual(results[1]["Name"], "Test2")

if __name__ == "__main__":
    unittest.main()
