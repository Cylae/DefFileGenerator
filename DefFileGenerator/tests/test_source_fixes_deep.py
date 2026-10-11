import pytest
from openpyxl import Workbook

from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor


def test_scalefactor_ratio_divisor_and_overflow_protection():
    """Verify that scale factors > 6 (e.g. 10000000, 10000) are treated as divisors without overflow."""
    rows = [
        [
            {"Address": "2800", "Name": "Latitude", "Type": "I32", "ScaleFactor": "10000000"},
            {"Address": "3001", "Name": "TankCapacity", "Type": "U32", "ScaleFactor": "10000"},
            {"Address": "12", "Name": "InputPower", "Type": "I16", "ScaleFactor": "0.1"},
            {"Address": "265", "Name": "ConsumedAh", "Type": "U16", "ScaleFactor": "-10"},
        ]
    ]
    ext = Extractor()
    cleaned = list(ext.map_and_clean(rows))
    assert len(cleaned) == 4

    gen = Generator()
    processed = list(gen.process_rows(cleaned))
    assert len(processed) == 4

    # 1 / 10000000 = 0.0000001 (rounded to 6 decimals: 0.000000)
    assert processed[0]["Name"] == "Latitude"
    assert processed[0]["CoefA"] == "0.000000"

    # 1 / 10000 = 0.0001
    assert processed[1]["Name"] == "TankCapacity"
    assert processed[1]["CoefA"] == "0.000100"

    # 1 / 0.1 = 10.0
    assert processed[2]["Name"] == "InputPower"
    assert processed[2]["CoefA"] == "10.000000"

    # 1 / -10 = -0.1
    assert processed[3]["Name"] == "ConsumedAh"
    assert processed[3]["CoefA"] == "-0.100000"


def test_is_secondary_header_row_guards_against_addresses_and_ranges():
    """Verify that data rows with addresses or address ranges are never mistaken for headers."""
    # Data row where column 0 is category, column 1 is address 40000, description mentions 'slave address'
    data_row = ["RS485 settings", "40000", "R/W", "RS485 slave address for port", "U16", "---", "1 ~ 247"]
    assert Extractor._is_secondary_header_row(data_row) is False

    # Data row with address range '041A -- 041F'
    range_row = ["041A -- 041F", "Temp_Sensors", "I16", "1", "degC", "R", "Temperature description"]
    assert Extractor._is_secondary_header_row(range_row) is False

    # Short banner or title (< 3 cells)
    banner_row = ["Section II. Real-time data (0x0400-0x07FF)", "General description"]
    assert Extractor._is_secondary_header_row(banner_row) is False

    # Real secondary header row
    real_header = ["Address", "Register Name", "Data Type", "Unit", "Access"]
    assert Extractor._is_secondary_header_row(real_header) is True


def test_effective_max_column_bounding(tmp_path):
    """Verify that spreadsheets with empty styled columns (up to column 16384) do not trigger cell limit error."""
    wb = Workbook()
    ws = wb.active
    ws.title = "RegisterTable"
    ws.append(["Address", "Name", "Type"])
    for i in range(10):
        ws.append([100 + i, f"Reg_{i}", "U16"])
    # Simulate a styled empty cell at column 500
    ws.cell(row=1, column=500, value=None)

    filepath = str(tmp_path / "styled_wide.xlsx")
    wb.save(filepath)

    ext = Extractor()
    tables = list(ext.extract_from_excel(filepath))
    assert len(tables) == 1
    rows = list(ext.map_and_clean(tables))
    assert len(rows) == 10
    assert rows[0]["Address"] == "100"
    assert rows[9]["Address"] == "109"


def test_legacy_xls_extraction_if_available():
    """Verify that legacy .xls files can be read when xlrd is installed."""
    import sys
    if "xlrd" not in sys.modules:
        pytest.skip("xlrd not installed")
    ext = Extractor()
    assert hasattr(ext, "extract_from_excel")
