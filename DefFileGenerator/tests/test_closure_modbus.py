import logging
import os
import tempfile
import unittest

from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator


class TestClosureModbus(unittest.TestCase):
    def setUp(self):
        logging.disable(logging.CRITICAL)
        self.g = Generator()
        self.tmpdir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.tmpdir.cleanup()
        logging.disable(logging.NOTSET)

    def test_overlapping_identical_ranges(self):
        out = os.path.join(self.tmpdir.name, "out.csv")
        cfg = GeneratorConfig(input_file=None, output=out, manufacturer="M", model="m")
        data = [
            {"Name": "V1", "Address": "100", "Type": "U32"},  # spans 100, 101
            {"Name": "V2", "Address": "100", "Type": "U32"},  # Identical
            {"Name": "V3", "Address": "100", "Type": "U16"},  # Contained
            {"Name": "V4", "Address": "101", "Type": "U16"},  # Adjoining and contained
        ]
        run_generator(cfg, iter(data))
        self.assertTrue(os.path.exists(out))
        report = self.g.validate_csv_detailed(out, strict=True)
        self.assertFalse(report.is_valid)
        issues = [i.message for i in report.issues if i.code == "ADDRESS_OVERLAP"]
        self.assertEqual(len(issues), 3)

    def test_bits_crossing_register_boundaries(self):
        self.assertFalse(self.g.validate_address("100_15_2", "BITS", strict=True))

        self.assertEqual(self.g.get_register_count("STRING", "100_5"), 3)
        self.assertEqual(self.g.get_register_count("STRING", "100_6"), 3)
        self.assertEqual(self.g.get_register_count("STRING", "100_7"), 4)

    def test_zero_length_and_unknown_types(self):
        self.assertEqual(self.g.get_register_count("STRING", "100_0"), 0)
        self.assertEqual(self.g.get_register_count("UNKNOWN_TYPE", "100"), 1)


if __name__ == "__main__":
    unittest.main()
