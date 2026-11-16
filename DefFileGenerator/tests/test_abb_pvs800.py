#!/usr/bin/env python3
"""
Apollo-Standard Test Suite for ABB PVS800-57B Inverter Definition.

Verifies:
  1. UTF-8-sig encoding and absence of decoding errors across all lines.
  2. Modbus Parameter Addressing Mode 3 (256 * Group + 2 * Index - 1).
  3. Word-swap integrity on ALL 32-bit types (both U32_W and I32_W).
  4. WebdynSunPM canonical telemetry and command tag assignments.
  5. Schema validation via Generator.validate_csv_detailed (strict=True, strict_overlap=False).
"""

from __future__ import annotations

import csv
import sys
from pathlib import Path

import pytest

# Ensure project root is available on sys.path
REPO_ROOT = Path(__file__).resolve().parent.parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from DefFileGenerator.def_gen import Generator  # noqa: E402
from DefFileGenerator.knowledge_miner import infer_sunpm_tag  # noqa: E402

REPO_DEF_PATH = REPO_ROOT / "Definitions" / "INV_ABB_PVS800-57B.csv"
REPO_DEF_2_PATH = REPO_ROOT / "Definitions" / "INV_ABB_PVS800-57B_2.csv"
DOWNLOADS_DEF_PATH = Path(r"c:\Users\alexandre.calais\Downloads\PVS800.csv")

DEF_PATHS = [REPO_DEF_PATH, REPO_DEF_2_PATH]


@pytest.mark.parametrize("def_path", DEF_PATHS)
def test_pvs800_definition_encoding_and_header(def_path: Path) -> None:
    """Verifies that the PVS800 definition file is encoded in valid UTF-8-sig."""
    assert def_path.is_file(), f"Missing definition file: {def_path}"

    with open(def_path, "rb") as f:
        raw_bytes = f.read()

    # Must start with UTF-8 BOM
    assert raw_bytes.startswith(b"\xef\xbb\xbf"), f"{def_path.name} must start with UTF-8 BOM"

    # Must decode cleanly without UnicodeDecodeError
    text = raw_bytes.decode("utf-8-sig")
    assert "\x00" not in text, f"{def_path.name} must not contain NUL bytes"

    # Header verification
    reader = csv.reader(text.splitlines(), delimiter=";")
    header = next(reader)
    assert len(header) >= 4
    assert header[0].lower() in {"modbus", "modbusrtu"}
    assert header[1] == "Inverter"
    assert header[2] == "ABB"
    assert "PVS800" in header[3]


@pytest.mark.parametrize("def_path", DEF_PATHS)
def test_pvs800_function_code_holding_registers(def_path: Path) -> None:
    """
    Verifies that all parameter rows use Function Code 3 (Read Holding Registers).
    ABB Modbus adapters (FSCA-01, FENA-01/11/21, RMBA-01) map all drive parameters
    and actual signals into Holding Registers and reject FC 04 with Exception 01.
    """
    with open(def_path, encoding="utf-8-sig") as f:
        reader = csv.reader(f, delimiter=";")
        next(reader)
        rows = [r for r in reader if r]

    for idx, row in enumerate(rows, start=2):
        fc = row[1].strip()
        assert fc == "3", (
            f"{def_path.name} Line {idx}: Expected Function Code 3 (Read Holding Registers), got {fc}"
        )


@pytest.mark.parametrize("def_path", DEF_PATHS)
def test_pvs800_word_swap_on_all_32bit_types(def_path: Path) -> None:
    """
    Verifies that all 32-bit registers (both unsigned and signed)
    have word-swap enabled (_W) per ABB PVS800 Modbus specifications.
    """
    with open(def_path, encoding="utf-8-sig") as f:
        reader = csv.reader(f, delimiter=";")
        next(reader)
        rows = [r for r in reader if r]

    assert len(rows) == 1053, f"Expected 1053 registers, found {len(rows)}"

    u32_count = 0
    i32_count = 0
    for idx, row in enumerate(rows, start=2):
        dtype = row[3].strip()
        # Must never have un-swapped U32 or I32
        assert dtype != "U32", f"{def_path.name} Line {idx}: U32 found without word-swap (_W)"
        assert dtype != "I32", f"{def_path.name} Line {idx}: I32 found without word-swap (_W)"

        if dtype == "U32_W":
            u32_count += 1
        elif dtype == "I32_W":
            i32_count += 1

    assert u32_count == 872, f"Expected 872 U32_W registers, got {u32_count}"
    assert i32_count == 131, f"Expected 131 I32_W registers, got {i32_count}"


@pytest.mark.parametrize("def_path", DEF_PATHS)
def test_pvs800_canonical_sunpm_tags(def_path: Path) -> None:
    """
    Verifies that essential inverter telemetry and command variables
    carry valid, unique WebdynSunPM canonical tags.
    """
    with open(def_path, encoding="utf-8-sig") as f:
        reader = csv.reader(f, delimiter=";")
        next(reader)
        rows = [r for r in reader if r]

    name_to_row = {r[5].strip(): r for r in rows}

    expected_tags = {
        "DC voltage": ("25857", "U32_W", "U_DC"),
        "Line current": ("25859", "U32_W", "I_AC"),
        "Frequency": ("25871", "U32_W", "FREQ"),
        "Grid voltage": ("25873", "U32_W", "U_AC_L1L2"),
        "Apparent power": ("25875", "I32_W", "S_AC"),
        "Power": ("25879", "I32_W", "P_AC"),
        "Reactive power": ("25883", "I32_W", "Q_AC"),
        "CosPhi": ("25887", "I32_W", "PF"),
        "Nominal power": ("25983", "U32_W", "P_RATED"),
        "Total energy": ("47115", "U32_W", "E_TOT"),
        "Daily kWh supplied": ("47131", "U32_W", "E_DAY"),
        "Ambient temperature": ("54017", "I32_W", "TEMP_AMB"),
        "Control board temperature": ("54019", "I32_W", "TEMP_INT"),
        "Insulation resistance": ("53817", "U32_W", "R_ISO"),
        "Tripping fault": ("26625", "U32_W", "FAULT_CODE"),
        "External power limit": ("33429", "U32_W", "cmdPwrPct"),
        "External limit ramp up": ("33475", "U32_W", "WMaxLimPct_RmpTms"),
    }

    seen_tags: set[str] = set()
    for name, (expected_addr, expected_type, expected_tag) in expected_tags.items():
        assert name in name_to_row, f"Variable '{name}' not found in {def_path.name}"
        row = name_to_row[name]
        addr = row[2].strip()
        dtype = row[3].strip()
        tag = row[6].strip()

        assert addr == expected_addr, (
            f"Mismatch for '{name}': expected addr {expected_addr}, got {addr}"
        )
        assert dtype == expected_type, (
            f"Mismatch for '{name}': expected type {expected_type}, got {dtype}"
        )
        assert tag == expected_tag, f"Mismatch for '{name}': expected tag {expected_tag}, got {tag}"

        assert tag not in seen_tags, f"Duplicate tag '{tag}' detected in {def_path.name}"
        seen_tags.add(tag)


@pytest.mark.parametrize("def_path", DEF_PATHS)
def test_pvs800_schema_validation(def_path: Path) -> None:
    """Verifies that the definition passes strict Webdyn schema validation."""
    gen = Generator()
    report = gen.validate_csv_detailed(
        str(def_path),
        strict=True,
        strict_overlap=False,
    )
    assert report.is_valid, (
        f"Validation issues in {def_path.name}: {[str(i) for i in report.issues]}"
    )
    assert report.register_count == 1053, (
        f"Expected 1053 registers in {def_path.name}, got {report.register_count}"
    )


def test_knowledge_miner_pvs800_tag_inference() -> None:
    """Verifies that the knowledge miner correctly infers SunPM tags for PVS800 variables."""
    assert infer_sunpm_tag("Line current", equipment_type="Inverter") == "I_AC"
    assert infer_sunpm_tag("DC voltage", equipment_type="Inverter") == "U_DC"
    assert infer_sunpm_tag("Grid voltage", equipment_type="Inverter") == "U_AC_L1L2"
    assert infer_sunpm_tag("Power", equipment_type="Inverter") == "P_AC"
    assert infer_sunpm_tag("Apparent power", equipment_type="Inverter") == "S_AC"
    assert infer_sunpm_tag("Reactive power", equipment_type="Inverter") == "Q_AC"
    assert infer_sunpm_tag("CosPhi", equipment_type="Inverter") == "PF"
    assert infer_sunpm_tag("Nominal power", equipment_type="Inverter") == "P_RATED"
    assert infer_sunpm_tag("Total energy", equipment_type="Inverter") == "E_TOT"
    assert infer_sunpm_tag("Daily kWh supplied", equipment_type="Inverter") == "E_DAY"
    assert infer_sunpm_tag("Ambient temperature", equipment_type="Inverter") == "TEMP_AMB"
    assert infer_sunpm_tag("Control board temperature", equipment_type="Inverter") == "TEMP_INT"
    assert infer_sunpm_tag("Insulation resistance", equipment_type="Inverter") == "R_ISO"
