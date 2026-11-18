"""
Test suite certifying the Huawei GENERIC Inverter Definition (INV_Huawei_GENERIC.csv)
and its synchronization across repository definition suites.
"""

import csv
from pathlib import Path

import pytest

from DefFileGenerator.def_gen import Generator

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
DEFINITIONS_DIR = REPO_ROOT / "Definitions"
TAGGED_DIR = REPO_ROOT / "tagged_definitions"
GENERATED_DIR = REPO_ROOT / "generated_definitions"


@pytest.fixture(scope="module")
def generator():
    return Generator()


def test_huawei_generic_files_exist():
    """Verify that INV_Huawei_GENERIC.csv exists in all canonical suites and OTH is purged."""
    def_file = DEFINITIONS_DIR / "INV_Huawei_GENERIC.csv"
    tagged_file = TAGGED_DIR / "INV_Huawei_GENERIC.csv"
    gen_file = GENERATED_DIR / "INV_Huawei_GENERIC.csv"

    assert def_file.is_file(), f"Missing {def_file}"
    assert tagged_file.is_file(), f"Missing {tagged_file}"
    assert gen_file.is_file(), f"Missing {gen_file}"

    # Verify misclassified OTH file is absent
    assert not (TAGGED_DIR / "OTH_Huawei_GENERIC.csv").exists(), (
        "OTH_Huawei_GENERIC.csv should be purged"
    )
    assert not (GENERATED_DIR / "OTH_Huawei_GENERIC.csv").exists(), (
        "OTH_Huawei_GENERIC.csv should be purged"
    )


@pytest.mark.parametrize("suite_dir", [DEFINITIONS_DIR, TAGGED_DIR, GENERATED_DIR])
def test_huawei_generic_encoding_and_schema(suite_dir):
    """Verify UTF-8 without BOM, 11-column CSV format, and exact header."""
    file_path = suite_dir / "INV_Huawei_GENERIC.csv"

    # 1. Zero BOM
    with open(file_path, "rb") as fp:
        raw_prefix = fp.read(4)
        assert not raw_prefix.startswith(b"\xef\xbb\xbf"), f"{file_path} contains UTF-8 BOM"

    # 2. Parse CSV
    with open(file_path, encoding="utf-8") as fp:
        reader = csv.reader(fp, delimiter=";")
        header = next(reader)
        rows = list(reader)

    # 3. Header check
    assert header[0] in ("Modbus", "modbusRTU"), f"Invalid protocol: {header[0]}"
    assert header[1] == "Inverter", f"Invalid category: {header[1]}"
    assert header[2] == "Huawei", f"Invalid manufacturer: {header[2]}"
    assert header[3] == "GENERIC", f"Invalid model: {header[3]}"
    assert len(header) == 11, f"Header columns {len(header)} != 11"

    # 4. Row integrity
    assert len(rows) == 256, f"Expected 256 registers, got {len(rows)}"
    for idx, r in enumerate(rows, 1):
        assert len(r) == 11, f"Row {idx} has {len(r)} columns instead of 11"
        assert r[0] == str(idx), f"Row {idx} index mismatch: {r[0]} != {idx}"
        assert r[1] in ("1", "2", "3", "4"), f"Row {idx} invalid function code: {r[1]}"
        assert r[2], f"Row {idx} empty address"
        assert r[3], f"Row {idx} empty type"
        assert r[5], f"Row {idx} empty name"


def test_huawei_generic_strict_validation(generator):
    """Verify that INV_Huawei_GENERIC.csv passes strict Webdyn validation with 0 issues."""
    file_path = DEFINITIONS_DIR / "INV_Huawei_GENERIC.csv"
    rep = generator.validate_csv_detailed(str(file_path), strict=True)
    assert rep.is_valid, f"Validation failed: {rep.issues}"
    assert len(rep.issues) == 0, f"Expected 0 issues, got {rep.issues}"
    assert rep.register_count == 256


def test_huawei_generic_sunpm_telemetry_tags():
    """Verify that telemetry and control tags are accurately assigned to real-time registers."""
    file_path = DEFINITIONS_DIR / "INV_Huawei_GENERIC.csv"
    with open(file_path, encoding="utf-8") as fp:
        reader = csv.reader(fp, delimiter=";")
        next(reader)
        rows = {r[2]: r for r in reader}

    # Core identification
    assert rows["30000_30"][6] == "MODEL"
    assert rows["30015_20"][6] == "SERIAL"
    assert rows["30073"][6] == "P_RATED"

    # Pmax and Qmax must not steal P_AC / Q_AC from instantaneous registers
    assert rows["30075"][6] == ""
    assert rows["30079"][6] == ""

    # Instantaneous AC Power
    assert rows["32080"][6] == "P_AC"
    assert rows["32080"][3] == "I32"
    assert rows["32080"][7] == "0.001000"
    assert rows["32080"][9] == "kW"

    assert rows["32082"][6] == "Q_AC"
    assert rows["32082"][3] == "I32"
    assert rows["32082"][7] == "0.001000"
    assert rows["32082"][9] == "kVar"

    # Three-phase currents
    assert rows["32072"][6] == "I_AC_1"
    assert rows["32074"][6] == "I_AC_2"
    assert rows["32076"][6] == "I_AC_3"

    # Three-phase voltages
    assert rows["32069"][6] == "U_AC_1"
    assert rows["32070"][6] == "U_AC_2"
    assert rows["32071"][6] == "U_AC_3"
    assert rows["32066"][6] == "U_AC_L1L2"
    assert rows["32067"][6] == "U_AC_L2L3"
    assert rows["32068"][6] == "U_AC_L3L1"

    # DC measurements
    assert rows["32064"][6] == "P_DC"
    assert rows["32016"][6] == "U_DC_STRING_1"
    assert rows["32017"][6] == "I_DC_STRING_1"

    # Energy yields
    assert rows["32106"][6] == "E_TOT"
    assert rows["32114"][6] == "E_DAY"

    # Remote Commands
    assert rows["40200"][6] == "CMD_ON"
    assert rows["40201"][6] == "CMD_OFF"
    assert rows["40122"][6] == "CMD_PF"
    assert rows["40125"][6] == "CMD_P_AC_SET_PERCENT"
    assert rows["40126"][6] == "CMD_P_AC_SET_ABS"


def test_huawei_v4_phase_current_tags():
    """Verify that Definitions/INV_Huawei_V4.csv has full 3-phase current tags."""
    v4_path = DEFINITIONS_DIR / "INV_Huawei_V4.csv"
    with open(v4_path, encoding="utf-8") as fp:
        reader = csv.reader(fp, delimiter=";")
        next(reader)
        rows = {r[2]: r for r in reader}

    assert rows["32072"][6] == "I_AC_1"
    assert rows["32074"][6] == "I_AC_2"
    assert rows["32076"][6] == "I_AC_3"
