"""
Targeted unit tests for Wave 24 coverage elevation across def_gen.py, extractor.py, main.py, and web/app.py.
"""

import sys
from unittest.mock import patch

import pytest

from DefFileGenerator.def_gen import (
    Generator,
    GeneratorConfig,
    WebdynDefConfig,
    run_generator,
)
from DefFileGenerator.def_gen import (
    main as def_gen_main,
)


def test_webdyn_def_config_none_paths():
    """Test WebdynDefConfig with None or string paths."""
    cfg = WebdynDefConfig(
        input_file=None,
        output_file=None,
        manufacturer="Mfg",
        model="Mod",
    )
    assert cfg.input_file is None
    assert cfg.output_file is None

    cfg2 = WebdynDefConfig(
        input_file="in.csv",
        output_file="out.csv",
        manufacturer="Mfg",
        model="Mod",
    )
    assert cfg2.input_file == "in.csv"
    assert cfg2.output_file == "out.csv"


def test_normalize_address_cached_edge_cases():
    """Test edge case branches in _normalize_address_cached."""
    gen = Generator()
    # Empty string or whitespace address part
    assert gen.normalize_address_val("   ") == ""
    # Invalid 0x prefix hex
    assert gen.normalize_address_val("0xZZ") == "0xZZ"
    # Invalid h suffix hex
    assert gen.normalize_address_val("ZZh") == "ZZh"
    # Negative 0x hex
    assert gen.normalize_address_val("-0x10") == "-16"
    # Negative h suffix hex
    assert gen.normalize_address_val("-10h") == "-16"


def test_sanitize_csv_field_control_chars_and_formulas():
    """Test sanitize_csv_field with non-printable characters and formula triggers."""
    gen = Generator()
    # String consisting entirely of unprintable control characters
    unprintable_only = "\x01\x02\x03"
    assert gen.sanitize_csv_field(unprintable_only) == ""

    # Formula starting with leading single space and negative number is sanitized with apostrophe
    assert gen.sanitize_csv_field(" -123") == "' -123"
    # Formula trigger with non-finite float or unparseable float
    assert gen.sanitize_csv_field("=ABC") == "'=ABC"
    assert gen.sanitize_csv_field("=inf") == "'=inf"


def test_normalize_type_str_variants():
    """Test normalize_type with bare 'str' and 'str<N>' patterns."""
    gen = Generator()
    assert gen.normalize_type("str") == "STRING"
    assert gen.normalize_type("str20") == "STR20"
    assert gen.normalize_type("str100") == "STR100"


def test_validate_address_exceptions():
    """Test validate_address exception paths."""
    gen = Generator()
    # Non-parseable base address in string representation
    assert not gen.validate_address("invalid_address_str", "U16")


def test_get_register_count_str_exception():
    """Test get_register_count STR<N> conversion error branches."""
    gen = Generator()
    # Address with unparseable non-integer length component
    assert gen.get_register_count("STRING", "40001_invalid") == 0


def test_determine_info1_direct_matches():
    """Test _determine_info1 with direct MODBUS_VALID_INFO1 strings and substrings."""
    gen = Generator()
    assert gen._determine_info1("1") == "1"
    assert gen._determine_info1("2") == "2"
    assert gen._determine_info1("3") == "3"
    assert gen._determine_info1("4") == "4"
    assert gen._determine_info1("read discrete input 0x02") == "2"
    assert gen._determine_info1("read input register 0x04") == "4"
    assert gen._determine_info1("read coil 0x01") == "1"
    assert gen._determine_info1("read holding register 0x03") == "3"


def test_calculate_coefficients_overflow_and_nan():
    """Test _calculate_coefficients with float overflow or non-finite numbers."""
    gen = Generator()
    # Extreme scale factor causing OverflowError
    coef_a, coef_b = gen._calculate_coefficients("1.0", "0.0", "1000")
    assert coef_a == "1.000000"
    assert coef_b == "0.000000"

    # NaN / Inf offset
    coef_a, coef_b = gen._calculate_coefficients("1.0", "nan", "0")
    assert coef_b == "0.000000"


def test_check_address_overlap_bits_edge_cases():
    """Test _check_address_overlap interval search with BITS slice comparisons."""
    gen = Generator()
    usage = {}
    warned = set()

    # Add first BITS entry at address 100_0_8
    gen._check_address_overlap("3", "100_0_8", "BITS", "var1", 2, usage, warned)

    # Add second non-overlapping BITS entry at address 100_8_8
    overlap1 = gen._check_address_overlap("3", "100_8_8", "BITS", "var2", 3, usage, warned)
    assert not overlap1

    # Add overlapping BITS entry at address 100_4_8
    overlap2 = gen._check_address_overlap("3", "100_4_8", "BITS", "var3", 4, usage, warned)
    assert overlap2


def test_run_generator_missing_input_file(tmp_path):
    """Test run_generator with non-existent input_file."""
    cfg = GeneratorConfig(
        input_file=str(tmp_path / "non_existent.csv"),
        output=str(tmp_path / "out.csv"),
    )
    assert not run_generator(cfg)

    # Missing both input_file and input_data
    cfg_empty = GeneratorConfig(output=str(tmp_path / "out.csv"))
    assert not run_generator(cfg_empty)


def test_def_gen_main_entrypoint():
    """Test CLI main entrypoint in def_gen.py."""
    with patch.object(sys, "argv", ["def_gen.py", "--help"]):
        with pytest.raises(SystemExit) as exc_info:
            def_gen_main()
        assert exc_info.value.code == 0
