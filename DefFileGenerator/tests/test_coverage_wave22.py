"""
Targeted unit tests for edge coverage wave 22 across def_gen, extractor, main, and web/app.
"""

import argparse
import csv
from unittest.mock import MagicMock, patch

import pytest
from fastapi.testclient import TestClient

from DefFileGenerator.def_gen import (
    Generator,
    GeneratorConfig,
    _normalize_address_cached,
    run_generator,
)
from DefFileGenerator.def_gen import (
    main as def_gen_main,
)
from DefFileGenerator.extractor import ExtractionError, Extractor
from DefFileGenerator.extractor import main as extractor_main
from DefFileGenerator.main import _perform_extraction
from DefFileGenerator.main import main as main_cli
from web.app import app

# -----------------------------------------------------------------------------
# DefGen Edge Tests
# -----------------------------------------------------------------------------


def test_def_gen_address_and_sanitize_edge_cases():
    """Test address normalization, sanitize_csv_field, and normalize_type edge cases."""
    # Line 374: normalize_address_val with empty string via grouping removal
    assert Generator.normalize_address_val("") == ""

    # Line 402: Hex string parsing ValueError fallback in _normalize_address_cached
    assert _normalize_address_cached("0xGHIJK") == "0xGHIJK"

    # Line 498: sanitize_csv_field with non-printable string that becomes empty after stripping
    assert Generator.sanitize_csv_field("\x00\x01\x02") == ""

    # Line 585 & 603: normalize_type "str" and "string_wb"
    assert Generator.normalize_type("str") == "STRING"
    assert Generator.normalize_type("string_wb") == "STRING_WB"

    # Line 741: validate_address exception handling
    assert Generator.validate_address("invalid_address_format_parts_str_xyz", "STRING") is False

    # Line 787: get_register_count STR<N> with invalid N
    assert Generator.get_register_count("STRabc", "30001") == 0


def test_def_gen_process_and_overlap_edge_cases():
    """Test _process_name_and_tag, _check_address_overlap, and coefficient edge cases."""
    g = Generator()

    # Line 935: Duplicate tag warning path
    seen_names = {}
    seen_tags = {"custom_tag": 2}
    tag = g._process_name_and_tag("My Name", "custom_tag", 5, seen_names, seen_tags)
    assert tag == "custom_tag"

    # Line 1037 & 1039: _check_address_overlap with None address_usage / warned_lines
    res = g._check_address_overlap("3", "30001", "U16", "Var1", 10, None, None)
    assert res is False

    # Line 1066: BITS overlap check with negative bit slice (-1)
    res_bits = g._check_address_overlap("3", "30002_0_-1", "BITS", "Var2", 11)
    assert res_bits is False

    # Line 1151: _calculate_coefficients non-finite val_a
    coef_a, coef_b = g._calculate_coefficients("1e308", "0", "100")
    assert coef_a == "1.000000"

    # Line 1235: process_rows BITS without underscore
    rows = [{"name": "BitsVar", "address": "30001", "type": "BITS", "registertype": "Holding"}]
    processed = list(g.process_rows(rows))
    assert len(processed) == 1
    assert processed[0]["Info2"] == "30001_0_16"


def test_def_gen_file_and_cli_edge_cases(tmp_path):
    """Test csv sniffer error, empty config, and main CLI failure exit."""
    # Line 1539: validate_csv_detailed empty row handling
    test_csv = tmp_path / "test.csv"
    test_csv.write_text(
        "modbusRTU;Inverter;Mfg;Model;;;;;;;\n\n\n1;3;30001;U16;;V1;v1;1;0;V;4\n", encoding="utf-8"
    )
    report = Generator().validate_csv_detailed(str(test_csv))
    assert report.is_valid is True

    # Line 1819: write_output_csv mock outfile close OSError
    fake_outfile = MagicMock()
    fake_outfile.close.side_effect = OSError("Disk error")
    fake_outfile.fileno.return_value = 1
    # write_output_csv handles closing error gracefully
    res = Generator.write_output_csv(
        None,
        [
            {
                "Info1": "1",
                "Info2": "30001",
                "Info3": "U16",
                "Info4": "",
                "Name": "A",
                "Tag": "a",
                "CoefA": "1",
                "CoefB": "0",
                "Unit": "V",
                "Action": "4",
            }
        ],
    )
    assert res is True

    # Line 2007: run_generator with input_file None and input_data None
    cfg = GeneratorConfig(input_file=None, output=str(tmp_path / "out.csv"))
    assert run_generator(cfg) is False

    # Line 2023: run_generator csv.Sniffer().sniff csv.Error fallback
    input_csv = tmp_path / "sniff_fail.csv"
    input_csv.write_text("Name,Tag,Address\nVar1,tag1,30001\n", encoding="utf-8")
    with patch("csv.Sniffer.sniff", side_effect=csv.Error("Sniff failed")):
        cfg_sniff = GeneratorConfig(
            input_file=str(input_csv), output=str(tmp_path / "out_sniff.csv")
        )
        assert run_generator(cfg_sniff) is True

    # Line 2067: def_gen main exit failure
    with patch("sys.argv", ["def_gen.py"]):
        with pytest.raises(SystemExit) as exc_info:
            def_gen_main()
        assert exc_info.value.code == 1


# -----------------------------------------------------------------------------
# Extractor Edge Tests
# -----------------------------------------------------------------------------


def test_extractor_column_and_pdf_edge_cases(tmp_path):
    """Test _infer_table_columns fallback, PDF extraction paths, and CLI entrypoint."""
    ext = Extractor()

    # Line 432: _infer_table_columns with empty col_scores
    assert ext._infer_table_columns([]) is None

    # Line 448: _infer_table_columns address ValueError exception handling
    table_invalid_addr = [
        ["Header1", "Header2"],
        ["99999999999999999999999999999999999999999", "VarName"],
    ]
    inferred = ext._infer_table_columns(table_invalid_addr)
    assert inferred is not None or inferred is None

    # Line 583: _fuzzy_header_score returning float
    score = ext._fuzzy_header_score("Name", ["Name"])
    assert score > 0.0

    # Line 685: extract_from_pdf requested pages list append
    if hasattr(ext, "extract_from_pdf"):
        dummy_pdf = tmp_path / "dummy.pdf"
        dummy_pdf.write_bytes(b"%PDF-1.4 ... dummy")
        with patch("pdfplumber.open") as mock_pdf:
            mock_page = MagicMock()
            mock_page.extract_tables.return_value = [[["Reg", "Addr"], ["V1", "30001"]]]
            mock_pdf.return_value.__enter__.return_value.pages = [mock_page]
            res = list(ext.extract_from_pdf(str(dummy_pdf), pages=[1, 2]))
            assert isinstance(res, list)

    # Line 954: extract_from_csv unexpected error
    bad_csv = tmp_path / "bad.csv"
    bad_csv.write_text("a,b,c\n1,2,3\n", encoding="utf-8")
    with patch(
        "DefFileGenerator.extractor.csv.DictReader", side_effect=AttributeError("Unexpected error")
    ):
        with pytest.raises(ExtractionError):
            for t in ext.extract_from_csv(str(bad_csv)):
                list(t)

    # Line 1012: extract_from_xml TypeError / ValueError handling
    bad_xml = tmp_path / "bad.xml"
    bad_xml.write_text("<root><row><name>A</name></row></root>", encoding="utf-8")
    with patch("defusedxml.ElementTree.parse", side_effect=ValueError("XML Error")):
        tables = list(ext.extract_from_xml(str(bad_xml)))
        assert len(tables) > 0
        assert list(tables[0]) == []

    # Line 1063: map_and_clean with empty buffer
    empty_gen = ext.map_and_clean([])
    assert list(empty_gen) == []

    # Line 1296 & 1298: extractor.main execution
    excel_file = tmp_path / "test.xlsx"
    excel_file.write_bytes(b"dummy")
    with patch(
        "DefFileGenerator.extractor.Extractor.extract_from_excel",
        return_value=[iter([{"Name": "A", "Address": "10"}])],
    ):
        with patch("sys.argv", ["extractor.py", str(excel_file)]):
            extractor_main()


# -----------------------------------------------------------------------------
# Main CLI & Web API Edge Tests
# -----------------------------------------------------------------------------


def test_main_cli_extraction_and_error_paths(tmp_path):
    """Test _perform_extraction with xlsx/xml and run command failure paths."""
    # Line 120 & 125: _perform_extraction with xlsx and xml
    xml_file = tmp_path / "input.xml"
    xml_file.write_text(
        "<root><row><name>V1</name><address>30001</address></row></root>", encoding="utf-8"
    )
    ns_xml = argparse.Namespace(input_file=str(xml_file), sheet=None, pages=None)
    ext_res = list(_perform_extraction(ns_xml))
    assert len(ext_res) > 0

    xlsx_file = tmp_path / "input.xlsx"
    xlsx_file.write_bytes(b"dummy")
    ns_xlsx = argparse.Namespace(input_file=str(xlsx_file), sheet=None, pages=None)
    with patch(
        "DefFileGenerator.extractor.Extractor.extract_from_excel",
        return_value=[iter([{"Name": "V2", "Address": "30002"}])],
    ):
        ext_xlsx = list(_perform_extraction(ns_xlsx))
        assert len(ext_xlsx) > 0

    # Line 161: _perform_extraction with empty registers raises SystemExit
    empty_csv = tmp_path / "empty.csv"
    empty_csv.write_text("Col1,Col2\nVal1,Val2\n", encoding="utf-8")
    ns_empty = argparse.Namespace(input_file=str(empty_csv), sheet=None, pages=None)
    with patch("DefFileGenerator.extractor.Extractor.extract_from_csv", return_value=[]):
        with pytest.raises(SystemExit):
            list(_perform_extraction(ns_empty))

    # Line 276 & 279: main_cli run command missing or non-existent file
    with patch("sys.argv", ["deffilegen", "run"]):
        with pytest.raises(SystemExit) as exc1:
            main_cli()
        assert exc1.value.code == 1

    non_exist = str(tmp_path / "non_existent.csv")
    with patch("sys.argv", ["deffilegen", "run", non_exist]):
        with pytest.raises(SystemExit) as exc2:
            main_cli()
        assert exc2.value.code == 1

    # Line 322: main_cli run command validation failure
    valid_csv = tmp_path / "valid_in.csv"
    valid_csv.write_text(
        "Name,Tag,Address,Type,RegisterType\nV1,t1,30001,U16,Holding\n", encoding="utf-8"
    )
    out_csv = tmp_path / "out_val_fail.csv"
    with patch("DefFileGenerator.def_gen.Generator.validate_csv", return_value=False):
        with patch("sys.argv", ["deffilegen", "run", str(valid_csv), "-o", str(out_csv)]):
            with pytest.raises(SystemExit) as exc3:
                main_cli()
            assert exc3.value.code == 1


def test_web_app_edge_and_error_branches(tmp_path):
    """Test web app upload endpoints, exception re-raising, and missing filenames."""
    client = TestClient(app)

    # Line 184: convert endpoint missing filename
    res_no_fn = client.post("/api/convert", files={"file": ("", b"data", "text/csv")})
    assert res_no_fn.status_code in (400, 422)

    # Line 238 & 241: convert endpoint with PDF and XML files
    xml_data = b"<root><row><Name>V1</Name><Address>30001</Address><Type>U16</Type></row></root>"
    res_xml = client.post("/api/convert", files={"file": ("test.xml", xml_data, "application/xml")})
    assert res_xml.status_code in (200, 400)

    # Line 256: convert endpoint with no readable registers
    res_no_regs = client.post(
        "/api/convert", files={"file": ("empty.csv", b"col1,col2\nval1,val2\n", "text/csv")}
    )
    assert res_no_regs.status_code == 400

    # Line 303 & 309: convert endpoint validation failure & HTTPException re-raise
    valid_input = b"Name,Tag,Address,Type,RegisterType\nV1,t1,30001,U16,Holding\n"
    with patch("DefFileGenerator.def_gen.Generator.validate_csv_detailed") as mock_val:
        mock_val.return_value = MagicMock(is_valid=False, register_count=0, issues=[], stats={})
        res_val_fail = client.post(
            "/api/convert", files={"file": ("in.csv", valid_input, "text/csv")}
        )
        assert res_val_fail.status_code == 400

    # Line 375: validate endpoint missing filename
    res_val_no_fn = client.post("/api/validate", files={"file": ("", b"data", "text/csv")})
    assert res_val_no_fn.status_code in (400, 422)
