#!/usr/bin/env python3
"""
Unit tests for PDF page range parsing and selection in Extractor.extract_from_pdf.
"""

from unittest.mock import MagicMock, mock_open, patch

from DefFileGenerator.extractor import Extractor


def create_mock_pdf(num_pages: int = 10):
    """Creates a mock pdfplumber PDF instance with specified number of pages."""
    mock_pdf = MagicMock()
    mock_pages = []
    for i in range(1, num_pages + 1):
        page = MagicMock()
        page.extract_tables.return_value = [
            [
                ["Address", "Name", "Type"],
                [f"3000{i}", f"Var_{i}", "U16"],
            ]
        ]
        mock_pages.append(page)
    mock_pdf.pages = mock_pages
    return mock_pdf


@patch("os.path.exists", return_value=True)
@patch("builtins.open", new_callable=mock_open, read_data=b"%PDF-1.4 dummy content")
@patch("pdfplumber.open")
def test_pdf_page_range_single_int(mock_pdfplumber_open, mock_file_open, mock_exists):
    """Test passing a single integer as pages."""
    mock_pdf = create_mock_pdf(10)
    mock_pdfplumber_open.return_value.__enter__.return_value = mock_pdf

    extractor = Extractor()
    tables = list(extractor.extract_from_pdf("dummy.pdf", pages=3))

    # Should only process page index 2 (1-based page 3)
    assert len(tables) == 1
    rows = list(tables[0])
    assert len(rows) == 1
    assert rows[0]["Address"] == "30003"


@patch("os.path.exists", return_value=True)
@patch("builtins.open", new_callable=mock_open, read_data=b"%PDF-1.4 dummy content")
@patch("pdfplumber.open")
def test_pdf_page_range_comma_string(mock_pdfplumber_open, mock_file_open, mock_exists):
    """Test passing comma-separated page string."""
    mock_pdf = create_mock_pdf(10)
    mock_pdfplumber_open.return_value.__enter__.return_value = mock_pdf

    extractor = Extractor()
    tables = list(extractor.extract_from_pdf("dummy.pdf", pages="1, 4, 7"))

    assert len(tables) == 3
    addresses = [list(t)[0]["Address"] for t in tables]
    assert addresses == ["30001", "30004", "30007"]


@patch("os.path.exists", return_value=True)
@patch("builtins.open", new_callable=mock_open, read_data=b"%PDF-1.4 dummy content")
@patch("pdfplumber.open")
def test_pdf_page_range_dash_string(mock_pdfplumber_open, mock_file_open, mock_exists):
    """Test passing dash-range string notation (e.g. '3-6')."""
    mock_pdf = create_mock_pdf(10)
    mock_pdfplumber_open.return_value.__enter__.return_value = mock_pdf

    extractor = Extractor()
    tables = list(extractor.extract_from_pdf("dummy.pdf", pages="3-6"))

    assert len(tables) == 4
    addresses = [list(t)[0]["Address"] for t in tables]
    assert addresses == ["30003", "30004", "30005", "30006"]


@patch("os.path.exists", return_value=True)
@patch("builtins.open", new_callable=mock_open, read_data=b"%PDF-1.4 dummy content")
@patch("pdfplumber.open")
def test_pdf_page_range_mixed_string(mock_pdfplumber_open, mock_file_open, mock_exists):
    """Test passing mixed comma and dash range string (e.g. '1, 3-5, 9')."""
    mock_pdf = create_mock_pdf(10)
    mock_pdfplumber_open.return_value.__enter__.return_value = mock_pdf

    extractor = Extractor()
    tables = list(extractor.extract_from_pdf("dummy.pdf", pages="1, 3-5, 9"))

    assert len(tables) == 5
    addresses = [list(t)[0]["Address"] for t in tables]
    assert addresses == ["30001", "30003", "30004", "30005", "30009"]


@patch("os.path.exists", return_value=True)
@patch("builtins.open", new_callable=mock_open, read_data=b"%PDF-1.4 dummy content")
@patch("pdfplumber.open")
def test_pdf_page_range_inverted_dash_string(mock_pdfplumber_open, mock_file_open, mock_exists):
    """Test passing inverted dash range notation (e.g. '5-3')."""
    mock_pdf = create_mock_pdf(10)
    mock_pdfplumber_open.return_value.__enter__.return_value = mock_pdf

    extractor = Extractor()
    tables = list(extractor.extract_from_pdf("dummy.pdf", pages="5-3"))

    assert len(tables) == 3
    addresses = [list(t)[0]["Address"] for t in tables]
    assert addresses == ["30005", "30004", "30003"]


@patch("os.path.exists", return_value=True)
@patch("builtins.open", new_callable=mock_open, read_data=b"%PDF-1.4 dummy content")
@patch("pdfplumber.open")
def test_pdf_page_range_out_of_bounds_and_invalid(
    mock_pdfplumber_open, mock_file_open, mock_exists
):
    """Test out-of-bound pages and invalid page strings emit warnings and process valid pages."""
    mock_pdf = create_mock_pdf(5)
    mock_pdfplumber_open.return_value.__enter__.return_value = mock_pdf

    extractor = Extractor()
    tables = list(extractor.extract_from_pdf("dummy.pdf", pages="1, 99, invalid, 2-4"))

    # Valid pages: 1, 2, 3, 4 (4 total)
    assert len(tables) == 4
    addresses = [list(t)[0]["Address"] for t in tables]
    assert addresses == ["30001", "30002", "30003", "30004"]
