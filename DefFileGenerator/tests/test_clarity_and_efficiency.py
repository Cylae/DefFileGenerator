#!/usr/bin/env python3
"""
Unit and benchmark tests for code clarity, efficiency optimizations,
constants, and developer documentation contracts.
"""

from __future__ import annotations

import inspect
import time
import unittest

from DefFileGenerator import def_gen, extractor, main
from DefFileGenerator.def_gen import (
    ACTION_READ_ONLY,
    ACTION_WRITE_ONLY,
    ALLOWED_ACTIONS,
    MAX_MODBUS_ADDRESS,
    MIN_MODBUS_ADDRESS,
    MODBUS_COIL,
    MODBUS_DISCRETE,
    MODBUS_HOLDING,
    MODBUS_INPUT,
    TYPE_SYNONYMS_COMPILED,
    CSVHeaderConfig,
    Generator,
    GeneratorConfig,
    RegisterEntry,
    ValidationIssue,
    ValidationReport,
    WebdynDefConfig,
)
from DefFileGenerator.extractor import (
    MODBUS_HEADER_KEYWORDS,
    NAME_INDICATOR_TERMS,
    NON_ADDRESS_COLUMN_TERMS,
    PDF_COMM_KEYWORDS,
    PDF_METADATA_KEYWORDS,
    Extractor,
)


class TestDomainConstants(unittest.TestCase):
    """Verifies that domain constants are correctly defined and typed."""

    def test_modbus_function_codes(self):
        self.assertEqual(MODBUS_COIL, "1")
        self.assertEqual(MODBUS_DISCRETE, "2")
        self.assertEqual(MODBUS_HOLDING, "3")
        self.assertEqual(MODBUS_INPUT, "4")

    def test_address_boundaries(self):
        self.assertEqual(MIN_MODBUS_ADDRESS, 0)
        self.assertEqual(MAX_MODBUS_ADDRESS, 65535)

    def test_action_constants(self):
        self.assertEqual(ACTION_WRITE_ONLY, "1")
        self.assertEqual(ACTION_READ_ONLY, "4")
        self.assertIn(ACTION_WRITE_ONLY, ALLOWED_ACTIONS)
        self.assertIn(ACTION_READ_ONLY, ALLOWED_ACTIONS)


class TestEfficiencyOptimizations(unittest.TestCase):
    """Verifies pre-compiled regular expressions and frozen keyword lookups."""

    def test_type_synonyms_compiled_tuple(self):
        self.assertIsInstance(TYPE_SYNONYMS_COMPILED, tuple)
        self.assertGreater(len(TYPE_SYNONYMS_COMPILED), 20)
        for pattern, replacement in TYPE_SYNONYMS_COMPILED:
            self.assertTrue(hasattr(pattern, "search"), "Pattern must be a compiled regex")
            self.assertIsInstance(replacement, str)

    def test_frozen_keyword_sets(self):
        self.assertIsInstance(MODBUS_HEADER_KEYWORDS, frozenset)
        self.assertIn("address", MODBUS_HEADER_KEYWORDS)
        self.assertIn("register", MODBUS_HEADER_KEYWORDS)

        self.assertIsInstance(PDF_METADATA_KEYWORDS, frozenset)
        self.assertIn("confidential", PDF_METADATA_KEYWORDS)

        self.assertIsInstance(PDF_COMM_KEYWORDS, frozenset)
        self.assertIn("baud rate", PDF_COMM_KEYWORDS)

        self.assertIsInstance(NAME_INDICATOR_TERMS, frozenset)
        self.assertIn("description", NAME_INDICATOR_TERMS)

        self.assertIsInstance(NON_ADDRESS_COLUMN_TERMS, frozenset)
        self.assertIn("country", NON_ADDRESS_COLUMN_TERMS)

    def test_normalize_type_benchmark(self):
        """Ensures that 1,000 type normalizations execute in under 50ms."""
        types_to_test = [
            "uint16",
            "float32",
            "int32",
            "uint64",
            "string 20",
            "bitfield16",
            "double",
            "unsigned short",
            "signed long",
            "float32 big endian",
        ]
        start = time.perf_counter()
        for _ in range(100):
            for t in types_to_test:
                Generator.normalize_type(t)
        elapsed = time.perf_counter() - start
        self.assertLess(elapsed, 0.1, f"Normalization took {elapsed:.4f}s, expected < 0.1s")


class TestDocumentationCompleteness(unittest.TestCase):
    """Ensures that all core classes and methods have informative docstrings."""

    def test_core_classes_have_docstrings(self):
        classes = [
            Generator,
            Extractor,
            WebdynDefConfig,
            CSVHeaderConfig,
            RegisterEntry,
            GeneratorConfig,
            ValidationIssue,
            ValidationReport,
        ]
        for cls in classes:
            doc = inspect.getdoc(cls)
            self.assertIsNotNone(doc, f"Class {cls.__name__} must have a docstring")
            self.assertGreater(len(doc.strip()), 10, f"Class {cls.__name__} docstring is too short")

    def test_public_methods_have_docstrings(self):
        gen_methods = [
            Generator.sanitize_csv_field,
            Generator.normalize_type,
            Generator.validate_type,
            Generator.normalize_address_val,
            Generator.validate_address,
            Generator.get_register_count,
            Generator.apply_address_offset,
            Generator.process_rows,
            Generator.validate_csv_detailed,
            Generator.validate_csv,
            Generator.write_output_csv,
        ]
        for m in gen_methods:
            doc = inspect.getdoc(m)
            self.assertIsNotNone(doc, f"Method {m.__name__} must have a docstring")

        extractor_methods = [
            Extractor.extract_from_excel,
            Extractor.extract_from_pdf,
            Extractor.extract_from_csv,
            Extractor.extract_from_xml,
            Extractor.map_and_clean,
        ]
        for m in extractor_methods:
            doc = inspect.getdoc(m)
            self.assertIsNotNone(doc, f"Method {m.__name__} must have a docstring")

    def test_modules_have_docstrings(self):
        for mod in (def_gen, extractor, main):
            doc = inspect.getdoc(mod)
            self.assertIsNotNone(doc, f"Module {mod.__name__} must have a docstring")
            self.assertGreater(
                len(doc.strip()), 20, f"Module {mod.__name__} docstring is too short"
            )


if __name__ == "__main__":
    unittest.main()
