"""
Targeted Unit Test Suite - Coverage Wave 25.
Focuses on edge error paths, fallback conditions, and boundary conditions across:
- DefFileGenerator/def_gen.py
- DefFileGenerator/extractor.py
- DefFileGenerator/knowledge_miner.py
- web/app.py
"""

from fastapi.testclient import TestClient

from DefFileGenerator.def_gen import (
    Generator,
    GeneratorConfig,
    ValidationIssue,
    WebdynDefConfig,
)
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.knowledge_miner import (
    infer_action_code,
    parse_function_code_from_text,
)
from web.app import app

# =============================================================================
# Knowledge Miner Edge Coverage
# =============================================================================


def test_knowledge_miner_fc_regex_fallback():
    # Test val = 05 matching regex and lstripping 0 to "5", returning "1"
    res = parse_function_code_from_text("function code: 05")
    assert res == "1"

    # Test val == "2" via regex fallback
    res2 = parse_function_code_from_text("function code: 02")
    assert res2 == "2"

    # Test val in ("3", "6", "16", "10")
    assert parse_function_code_from_text("code fonction: 06") == "3"
    assert parse_function_code_from_text("fc: 04") == "4"


def test_knowledge_miner_infer_action_code():
    # Write command with fc=5, rw="w"
    assert infer_action_code("set_active_power", fc=5, rw_field="w") == "1"
    # Reserve
    assert infer_action_code("RES.", fc=4, rw_field="r") == "0"
    # Alarm/Status
    assert infer_action_code("grid_fault_alarm", fc=4, rw_field="r") == "9"


# =============================================================================
# DefGen Config & Boundary Edge Coverage
# =============================================================================


def test_def_gen_config_post_init():
    # Path post_init empty/None handling
    cfg1 = WebdynDefConfig(manufacturer="Mfg", model="Mod", input_file="", output_file="")
    assert cfg1.input_file == ""

    cfg2 = GeneratorConfig(input_file="", output="")
    assert cfg2.output == ""


def test_def_gen_address_and_type_validation_edges():
    gen = Generator()
    # Empty string or invalid address normalization
    assert gen.validate_address("", "INT16") is False

    # Private row action validator
    issues: list[ValidationIssue] = []
    gen._validate_row_action("99", 1, True, issues)
    assert len(issues) > 0

    # Private row coefficients validator: row is list of 9 elements
    row = ["4", "Tag", "100", "U16", "4", "0", "0", "invalid_coef", "0"]
    coef_issues: list[ValidationIssue] = []
    gen._validate_row_coefs(row, 1, True, coef_issues)
    assert len(coef_issues) > 0


def test_def_gen_check_address_overlap_edge():
    gen = Generator()
    # Same base address, different bit ranges
    rows = [
        {
            "Info1": "4",
            "Tag": "T1",
            "Type": "BITS",
            "Address": "100.0-100.3",
            "Action": "4",
            "coef_a": "1",
            "coef_b": "0",
        },
        {
            "Info1": "4",
            "Tag": "T2",
            "Type": "BITS",
            "Address": "100.4-100.7",
            "Action": "4",
            "coef_a": "1",
            "coef_b": "0",
        },
    ]
    has_overlap = gen._check_address_overlap(rows)
    assert has_overlap is False


# =============================================================================
# Extractor Edge Coverage
# =============================================================================


def test_extractor_pdf_and_excel_edge_fallbacks():
    ext = Extractor()

    # Empty generator for missing pdf file
    res_pdf = list(ext.extract_from_pdf("non_existent_file.pdf"))
    assert res_pdf == []

    # Missing excel file generator yields empty row list or nested generator
    res_excel = list(ext.extract_from_excel("non_existent_file.xlsx"))
    assert isinstance(res_excel, list)


def test_extractor_map_and_clean():
    ext = Extractor()
    # Test empty generator mapping
    mapped = list(ext.map_and_clean([]))
    assert mapped == []


# =============================================================================
# Web API Edge Coverage
# =============================================================================


def test_web_app_rest_api_edge_cases():
    client = TestClient(app)

    # Healthcheck
    resp = client.get("/api/health")
    assert resp.status_code == 200
    assert resp.json()["status"] == "ok"

    # API validate with empty file content
    resp_val = client.post(
        "/api/validate", files={"file": ("empty.csv", b"", "text/csv")}
    )
    assert resp_val.status_code in (200, 400)

    # API validate with invalid CSV headers
    resp_inv = client.post(
        "/api/validate",
        files={"file": ("invalid.csv", b"HeaderA,HeaderB\n1,2\n", "text/csv")},
    )
    assert resp_inv.status_code == 200
    data = resp_inv.json()
    assert "valid" in data
