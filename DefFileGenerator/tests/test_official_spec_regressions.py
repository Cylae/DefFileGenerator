"""Regression tests derived from the official WebdynSunPM definition-file specification."""

import unittest

from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor


class TestOfficialDefinitionSpecification(unittest.TestCase):
    def test_action_10_constant_is_supported(self):
        generator = Generator()

        self.assertIn("10", generator.allowed_actions)

    def test_raw_requires_positive_even_byte_length(self):
        self.assertTrue(Generator.validate_type("RAW"))
        self.assertTrue(Generator.validate_address("40000_10", "RAW"))
        self.assertEqual(Generator.get_register_count("RAW", "40000_10"), 5)
        self.assertFalse(Generator.validate_address("40000", "RAW"))
        self.assertFalse(Generator.validate_address("40000_3", "RAW"))

    def test_fragmented_reversed_pdf_headers_are_recognized(self):
        extractor = Extractor()

        self.assertTrue(extractor._fuzzy_header_matches("[1] s s e r d d a", ["address"]))
        self.assertTrue(extractor._fuzzy_header_matches("e p y t a at D", ["data type"]))
        self.assertTrue(extractor._fuzzy_header_matches("n o pti ri c s e d", ["description"]))
        self.assertTrue(extractor._fuzzy_header_matches("r e st gi e R", ["register"]))
        self.assertFalse(extractor._fuzzy_header_matches("Default value", ["address"]))

    def test_reversed_headers_map_a_complete_register_row(self):
        extractor = Extractor()
        tables = [
            [
                {
                    "[1] s s e r d d a": "40001",
                    "r e st gi e R": "Grid voltage",
                    "e p y t a at D": "Unsigned 16",
                    "nit U": "V",
                }
            ]
        ]

        mapped = list(extractor.map_and_clean(tables))

        self.assertEqual(len(mapped), 1)
        self.assertEqual(mapped[0]["Address"], "40001")
        self.assertEqual(mapped[0]["Name"], "Grid voltage")
        self.assertEqual(mapped[0]["Type"], "U16")
        self.assertEqual(mapped[0]["Unit"], "V")

    def test_voltage_range_is_not_misidentified_as_tag(self):
        rows = list(
            Extractor().map_and_clean(
                [
                    [
                        {
                            "Address": "2433",
                            "Description": "Fixed power factor",
                            "Data type": "I16",
                            "Voltage range": "[-10000,10000]",
                        }
                    ]
                ]
            )
        )

        self.assertEqual(len(rows), 1)
        self.assertNotIn("Tag", rows[0])


if __name__ == "__main__":
    unittest.main()
