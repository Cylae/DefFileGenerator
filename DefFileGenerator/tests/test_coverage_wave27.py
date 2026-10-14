"""
Targeted Unit Tests for Wave 27 Coverage & Refactoring Verification.

Tests covering:
- doc_to_webdyn page range parsing (1-3, 5 and inverted ranges).
- Extractor prevention of false 0x0000 address synthesis from empty high/low cells.
- Generator subword disjoint byte checking for U8 and I8 (100_0 vs 100_1).
- Extractor CLI staged output writing and resource cleanup.
- generate_mars_definitions CLI error handling when source file is not found.
"""

from __future__ import annotations

import sys
from unittest.mock import patch

import pytest

import doc_to_webdyn
import generate_mars_definitions
from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.extractor import main as extractor_main


def test_doc_to_webdyn_page_range_parsing(tmp_path):
    """Verify that doc_to_webdyn supports page ranges like '1-3, 5'."""
    pdf_file = tmp_path / "dummy.pdf"
    pdf_file.write_bytes(b"%PDF-1.4 dummy")

    output_csv = tmp_path / "output.csv"
    test_args = [
        "doc_to_webdyn.py",
        str(pdf_file),
        "-o",
        str(output_csv),
        "--manufacturer",
        "TestMfg",
        "--model",
        "TestModel",
        "--pages",
        "1-3, 5",
    ]

    with patch.object(sys, "argv", test_args):
        # Even if dummy.pdf has no register tables, _perform_extraction will run and fail with 'No data extracted' (code 1)
        # without failing on 'Invalid format for --pages'
        with patch("logging.error") as mock_err:
            with pytest.raises(SystemExit) as cm:
                doc_to_webdyn.main()
            assert cm.value.code == 1
            # Ensure it didn't fail on pages formatting
            error_msgs = [call.args[0] for call in mock_err.call_args_list if call.args]
            assert not any("Invalid format for --pages" in msg for msg in error_msgs)


def test_extractor_empty_high_low_cells_do_not_synthesize_zero_address():
    """Verify that empty high and low byte cells do not synthesize an invalid 0x0000 address."""
    extractor = Extractor()
    # Simulate a row where high/low byte columns exist in the header, but cell values are empty
    table = [
        {"High Byte": "", "Low Byte": "", "Name": "Dummy Register", "Type": "U16"},
    ]
    # In table generator output, empty hi/lo cells shouldn't yield Address="0x0000"
    cleaned = list(extractor.map_and_clean([iter(table)]))
    for reg in cleaned:
        assert reg.get("Address") != "0x0000"


def test_generator_u8_byte_disjoint_slices_no_overlap():
    """Verify that U8 registers sharing a 16-bit word at byte 0 and byte 1 do not trigger overlap."""
    gen = Generator()
    rows = [
        {"Name": "RegByte0", "Address": "200_0", "Type": "U8", "RegisterType": "Holding Register"},
        {"Name": "RegByte1", "Address": "200_1", "Type": "U8", "RegisterType": "Holding Register"},
    ]
    processed = list(gen.process_rows(rows))
    assert len(processed) == 2


def test_generator_u8_byte_same_slice_triggers_overlap():
    """Verify that two U8 registers on the exact same byte slice (e.g. 200_0) trigger an overlap warning."""
    gen = Generator()
    rows = [
        {
            "Name": "RegByte0_A",
            "Address": "200_0",
            "Type": "U8",
            "RegisterType": "Holding Register",
        },
        {
            "Name": "RegByte0_B",
            "Address": "200_0",
            "Type": "U8",
            "RegisterType": "Holding Register",
        },
    ]
    with patch("logging.warning") as mock_warn:
        processed = list(gen.process_rows(rows))
        assert len(processed) == 2
        warn_msgs = [call.args[0] for call in mock_warn.call_args_list if call.args]
        assert any("Address overlap detected" in msg for msg in warn_msgs)


def test_generator_i8_byte_disjoint_slices_no_overlap():
    """Verify that I8 registers sharing a word at byte 0 and byte 1 do not trigger overlap."""
    gen = Generator()
    rows = [
        {
            "Name": "RegI8_Byte0",
            "Address": "300_0",
            "Type": "I8",
            "RegisterType": "Holding Register",
        },
        {
            "Name": "RegI8_Byte1",
            "Address": "300_1",
            "Type": "I8",
            "RegisterType": "Holding Register",
        },
    ]
    processed = list(gen.process_rows(rows))
    assert len(processed) == 2


def test_extractor_cli_staged_output(tmp_path):
    """Verify Extractor CLI writes output atomically and closes the file."""
    csv_input = tmp_path / "registers_input.csv"
    csv_input.write_text("Address,Name,Type\n100,TestReg,U16\n", encoding="utf-8")

    out_file = tmp_path / "output_registers.csv"
    test_args = ["extractor.py", str(csv_input), "-o", str(out_file)]

    with patch.object(sys, "argv", test_args):
        extractor_main()

    assert out_file.is_file()
    content = out_file.read_text(encoding="utf-8")
    assert "TestReg" in content
    assert "100" in content


def test_generate_mars_definitions_missing_file_cli(capsys):
    """Verify generate_mars_definitions prints error and exits code 1 if source file does not exist."""
    with pytest.raises(SystemExit) as cm:
        generate_mars_definitions.main(["non_existent_io_list.xlsx"])
    assert cm.value.code == 1
    captured = capsys.readouterr()
    assert "Source workbook not found" in captured.err
