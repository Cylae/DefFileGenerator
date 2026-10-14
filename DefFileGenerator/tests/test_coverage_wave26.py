"""
Targeted Unit Tests for Wave 26 Coverage Expansion.

Covering edge branches in DefFileGenerator/extractor.py and DefFileGenerator/def_gen.py:
- Zip member size limits in Excel extraction.
- Legacy XLS / BadZipFile fallback handling.
- PDF page range parsing (including inverted ranges e.g. '5-2').
- PDF text fallback extraction table generation.
- CSV NUL character detection and encoding fallback.
- XML security and parse exception paths.
- Extractor stage 4 fuzzy header matching and HighByte/LowByte combinations.
- ScaleFactor float conversions in map_and_clean.
- Generator _calculate_coefficients extreme scale factor / non-finite float handling.
"""

from __future__ import annotations

import zipfile
from unittest.mock import MagicMock, patch

import pytest

from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import ExtractionError, Extractor


def test_excel_zip_member_size_limit(tmp_path):
    """Test ExtractionError raised when Excel zip member exceeds size limit."""
    excel_file = tmp_path / "huge.xlsx"
    with zipfile.ZipFile(excel_file, "w") as zf:
        zf.writestr("xl/worksheets/sheet1.xml", b"x" * 100)

    extractor = Extractor()
    with patch("zipfile.ZipFile") as mock_zip:
        mock_archive = MagicMock()
        mock_member = MagicMock()
        mock_member.file_size = 100 * 1024 * 1024  # > 64MB limit
        mock_archive.infolist.return_value = [mock_member]
        mock_zip.return_value.__enter__.return_value = mock_archive

        with pytest.raises(ExtractionError, match="archive member size"):
            list(extractor.extract_from_excel(str(excel_file)))


def test_excel_zip_archive_total_size_limit(tmp_path):
    """Test ExtractionError raised when Excel archive total uncompressed size exceeds limit."""
    excel_file = tmp_path / "huge_total.xlsx"
    with zipfile.ZipFile(excel_file, "w") as zf:
        zf.writestr("xl/worksheets/sheet1.xml", b"x" * 100)

    extractor = Extractor()
    with patch("zipfile.ZipFile") as mock_zip:
        mock_archive = MagicMock()
        # 3 members of 50MB each = 150MB total (> 128MB total limit, but each < 64MB member limit)
        mock_members = [MagicMock(file_size=50 * 1024 * 1024) for _ in range(3)]
        mock_archive.infolist.return_value = mock_members
        mock_zip.return_value.__enter__.return_value = mock_archive

        with pytest.raises(ExtractionError, match="maximum expanded archive size"):
            list(extractor.extract_from_excel(str(excel_file)))


def test_excel_bad_zip_file_fallback_to_xlrd(tmp_path):
    """Test fallback to xlrd when openpyxl encounters BadZipFile on legacy file."""
    fake_xls = tmp_path / "legacy.xls"
    fake_xls.write_bytes(b"not a zip file")

    extractor = Extractor()
    with patch("DefFileGenerator.extractor.xlrd") as mock_xlrd:
        mock_wb = MagicMock()
        mock_wb.sheet_names.return_value = ["Sheet1"]
        mock_ws = MagicMock()
        mock_ws.nrows = 2
        mock_ws.ncols = 2
        mock_ws.cell_value.side_effect = lambda r, c: f"val_{r}_{c}"
        mock_wb.sheet_by_name.return_value = mock_ws
        mock_xlrd.open_workbook.return_value = mock_wb

        sheets = list(extractor.extract_from_excel(str(fake_xls)))
        assert len(sheets) == 1
        rows = list(sheets[0])
        assert len(rows) > 0


def test_pdf_inverted_page_range_parsing(tmp_path):
    """Test PDF page range parsing with inverted ranges like '5-2'."""
    pdf_path = tmp_path / "dummy.pdf"
    pdf_path.write_bytes(b"%PDF-1.4 dummy")

    extractor = Extractor()
    with patch("pdfplumber.open") as mock_pdf_open:
        mock_pdf = MagicMock()
        pages_mock = [MagicMock() for _ in range(10)]
        for p in pages_mock:
            p.extract_tables.return_value = [[["Address", "Type"], ["100", "U16"]]]
        mock_pdf.pages = pages_mock
        mock_pdf_open.return_value.__enter__.return_value = mock_pdf

        tables = list(extractor.extract_from_pdf(str(pdf_path), pages="5-3"))
        assert len(tables) == 3


def test_pdf_text_based_table_fallback(tmp_path):
    """Test text regex line extraction when no PDF tables are extracted."""
    pdf_path = tmp_path / "text.pdf"
    pdf_path.write_bytes(b"%PDF-1.4 text")

    extractor = Extractor()
    with patch("pdfplumber.open") as mock_pdf_open:
        mock_pdf = MagicMock()
        page_mock = MagicMock()
        page_mock.extract_tables.return_value = []
        page_mock.extract_text.return_value = "40001 2 U16 ActivePower\n40003 1 U16 ReactivePower"
        mock_pdf.pages = [page_mock]
        mock_pdf_open.return_value.__enter__.return_value = mock_pdf

        tables = list(extractor.extract_from_pdf(str(pdf_path)))
        assert len(tables) == 1
        rows = list(tables[0])
        assert len(rows) == 2
        assert rows[0]["Address"] == "40001"
        assert rows[0]["Type"] == "U16"
        assert rows[0]["Name"] == "ActivePower"


def test_csv_nul_character_rejection(tmp_path):
    """Test CSV file with NUL character raises ExtractionError."""
    csv_file = tmp_path / "nul.csv"
    csv_file.write_bytes(b"Address;Type\x00;Name\n40001;U16;Test")

    extractor = Extractor()
    tables = list(extractor.extract_from_csv(str(csv_file)))
    with pytest.raises(ExtractionError, match="NUL characters"):
        list(tables[0])


def test_highbyte_lowbyte_combination():
    """Test map_and_clean synthesizing Address from HighByte and LowByte columns."""
    extractor = Extractor()
    raw_data = [
        [
            {
                "HighByte": "12",
                "LowByte": "34",
                "Name": "CombinedVar",
                "Type": "U16",
            }
        ]
    ]

    mapped = list(extractor.map_and_clean(raw_data))
    assert len(mapped) == 1
    # 0x1234 normalized to decimal is 4660
    assert mapped[0]["Address"] == "4660"
    assert mapped[0]["Name"] == "CombinedVar"


def test_scalefactor_float_conversion():
    """Test ScaleFactor parsing in map_and_clean with float values."""
    extractor = Extractor()
    raw_data = [
        [
            {
                "Address": "40001",
                "Name": "ScaledVar",
                "Type": "U16",
                "ScaleFactor": "2",
            },
            {
                "Address": "40002",
                "Name": "InvertedGainVar",
                "Type": "U16",
                "ScaleFactor": "0.1",
            },
        ]
    ]

    mapped = list(extractor.map_and_clean(raw_data))
    assert len(mapped) == 2
    assert mapped[0]["Factor"] == "100.0"
    assert mapped[1]["Factor"] == "10.0"


def test_calculate_coefficients_extreme_scalefactor():
    """Test Generator._calculate_coefficients with overflow/extreme ScaleFactor."""
    coef_a, coef_b = Generator._calculate_coefficients("1.0", "0.0", "1000")
    assert coef_a == "1.000000"
    assert coef_b == "0.000000"

    coef_a_nan, coef_b_nan = Generator._calculate_coefficients("NaN", "Inf", "0")
    assert coef_a_nan == "1.000000"
    assert coef_b_nan == "0.000000"
