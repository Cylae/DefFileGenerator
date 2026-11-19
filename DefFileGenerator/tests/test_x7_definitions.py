#!/usr/bin/env python3
"""
Apollo-Standard Test Suite for X7 Commercial & Industrial BESS Definitions.
Validates:
1. 5 Apollo Pillars compliance on all 7 generated definitions
2. 100% byte-identical redundant synchronization across directories
3. Exhaustive register mapping, types, scaling factors, units, and canonical SunPM tags
4. Cell instrumentation expansion (128 temperatures + 240 voltages)
5. Zero address collisions and zero duplicate single-tags
6. Manifest indexing and source document integrity
"""

import csv
import hashlib
import json
from pathlib import Path

from DefFileGenerator.production_governor import verify_file_pillars
from DefFileGenerator.x7_generator import parse_x7_workbook

WORKSPACE_DIR = Path(__file__).resolve().parent.parent.parent

X7_FILENAMES = [
    "BAT_X7_BESS-AllInOne.csv",
    "BAT_X7_EMS.csv",
    "BAT_X7_BMS.csv",
    "BAT_X7_BMS-Cells.csv",
    "MET_X7_Meters.csv",
    "INV_X7_PCS.csv",
    "OTH_X7_Cabinet.csv",
]


def test_x7_source_document_and_manifest():
    """Verify source workbook exists, hash matches, and manifest is properly indexed."""
    src_file = (
        WORKSPACE_DIR
        / "sources"
        / "manuals"
        / "X7"
        / "X7-ModbusTcp-PORT502-20261105-V1.0.6-Beta.xlsx"
    )
    assert src_file.is_file(), f"Source document missing: {src_file}"

    data = src_file.read_bytes()
    assert len(data) == 57075, f"Unexpected source size: {len(data)}"
    sha = hashlib.sha256(data).hexdigest()
    assert sha == "17630570efc930b43e01e76a7679f653d3b8fc6c32f82cb0887f76341351dae2"

    manifest_path = WORKSPACE_DIR / "sources" / "MANIFEST.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    x7_docs = [d for d in manifest["documents"] if d.get("document_id") == "SRC-X7-001"]
    assert len(x7_docs) == 1, "SRC-X7-001 not found in MANIFEST.json"
    doc = x7_docs[0]
    assert doc["manufacturer"] == "X7"
    assert doc["sha256"] == sha
    assert len(doc["definitions_generated"]) == 7


def test_x7_production_pillars_all_files():
    """Verify that all 7 X7 definitions in production/ comply strictly with 5 Apollo Pillars."""
    prod_dir = WORKSPACE_DIR / "production"
    for fname in X7_FILENAMES:
        fpath = prod_dir / fname
        assert fpath.is_file(), f"Production file missing: {fpath}"
        is_valid, issues = verify_file_pillars(fpath)
        assert is_valid, f"Pillar check failed for {fname}: {issues}"

        # Additional UTF-8 without BOM verification
        raw_prefix = fpath.read_bytes()[:4]
        assert not raw_prefix.startswith(b"\xef\xbb\xbf"), f"BOM detected in {fname}"


def test_x7_redundancy_synchronization():
    """Verify byte-identical synchronization across all reference directories."""
    folders = [
        WORKSPACE_DIR / "Definitions" / "X7",
        WORKSPACE_DIR / "X7_Definitions",
        WORKSPACE_DIR / "generated_definitions",
        WORKSPACE_DIR / "tagged_definitions",
    ]

    for fname in X7_FILENAMES:
        prod_bytes = (WORKSPACE_DIR / "production" / fname).read_bytes()
        for fld in folders:
            target = fld / fname
            assert target.is_file(), f"Redundant replica missing: {target}"
            assert target.read_bytes() == prod_bytes, f"Byte mismatch in {target}"


def test_x7_all_in_one_structure_and_completeness():
    """Verify master All-In-One definition has 794 registers, sequential index, and unique keys."""
    aio_path = WORKSPACE_DIR / "production" / "BAT_X7_BESS-AllInOne.csv"
    lines = aio_path.read_text(encoding="utf-8").splitlines()
    assert len(lines) == 795, f"Expected 795 lines (1 header + 794 regs), got {len(lines)}"

    with open(aio_path, encoding="utf-8") as fp:
        reader = list(csv.reader(fp, delimiter=";"))

    header = reader[0]
    assert header[0] == "Modbus"
    assert header[1] == "Battery"
    assert header[2] == "X7"
    assert header[3] == "BESS-AllInOne"
    assert len(header) == 11

    rows = reader[1:]
    assert len(rows) == 794

    seen_fc_addr = set()
    for idx, r in enumerate(rows, start=1):
        assert r[0] == str(idx), f"Index mismatch at row {idx}: {r[0]}"
        assert len(r) == 11, f"Row {idx} has {len(r)} columns instead of 11"
        fc = r[1]
        addr = r[2]
        key = (fc, addr)
        assert key not in seen_fc_addr, f"Duplicate (FC, Addr) {key} at row {idx}"
        seen_fc_addr.add(key)


def test_x7_cell_instrumentation_precision():
    """Verify detailed individual cell voltage (240) and temperature (128) expansion."""
    cells_path = WORKSPACE_DIR / "production" / "BAT_X7_BMS-Cells.csv"
    with open(cells_path, encoding="utf-8") as fp:
        reader = list(csv.reader(fp, delimiter=";"))

    rows = reader[1:]
    assert len(rows) == 368, f"Expected 368 cell registers, got {len(rows)}"

    reg_map = {r[2]: r for r in rows}

    # Verify 128 temperatures: 6200..6327, INT16 (I16), scale 1.0, unit °C
    for i in range(128):
        addr = str(6200 + i)
        assert addr in reg_map, f"Missing temp register {addr}"
        r = reg_map[addr]
        assert r[1] == "4", f"Expected FC 4 for temp {addr}"
        assert r[3] == "I16", f"Expected I16 for temp {addr}, got {r[3]}"
        assert r[6] == f"T_CELL_{i + 1}", f"Expected tag T_CELL_{i + 1}, got {r[6]}"
        assert float(r[7]) == 1.0, f"Expected CoefA 1.0, got {r[7]}"
        assert r[9] == "°C", f"Expected unit °C, got {r[9]}"

    # Verify 240 voltages: 6600..6839, UINT16 (U16), scale 0.001, unit V
    for i in range(240):
        addr = str(6600 + i)
        assert addr in reg_map, f"Missing voltage register {addr}"
        r = reg_map[addr]
        assert r[1] == "4", f"Expected FC 4 for voltage {addr}"
        assert r[3] == "U16", f"Expected U16 for voltage {addr}, got {r[3]}"
        assert r[6] == f"U_CELL_{i + 1}", f"Expected tag U_CELL_{i + 1}, got {r[6]}"
        assert float(r[7]) == 0.001, f"Expected CoefA 0.001, got {r[7]}"
        assert r[9] == "V", f"Expected unit V, got {r[9]}"


def test_x7_safety_alarms_and_controls():
    """Verify cabinet safety alarms (FC 2) and control setpoints (FC 3)."""
    cab_path = WORKSPACE_DIR / "production" / "OTH_X7_Cabinet.csv"
    with open(cab_path, encoding="utf-8") as fp:
        cab_rows = {r[2]: r for r in list(csv.reader(fp, delimiter=";"))[1:]}

    # Emergency stop feedback
    assert "54200" in cab_rows
    assert cab_rows["54200"][1] == "2"
    assert cab_rows["54200"][10] == "9"  # Action 9 = Alarm

    # Fire extinguisher feedback
    assert "54203" in cab_rows
    assert cab_rows["54203"][1] == "2"
    assert cab_rows["54203"][10] == "9"

    # Combustible gas alarm
    assert "54211" in cab_rows
    assert cab_rows["54211"][1] == "2"
    assert cab_rows["54211"][10] == "9"

    # EMS Controls in BAT_X7_EMS.csv
    ems_path = WORKSPACE_DIR / "production" / "BAT_X7_EMS.csv"
    with open(ems_path, encoding="utf-8") as fp:
        ems_rows = {r[2]: r for r in list(csv.reader(fp, delimiter=";"))[1:]}

    assert "2000" in ems_rows
    assert ems_rows["2000"][1] == "3"
    assert ems_rows["2000"][10] == "1"  # Action 1 = Command

    assert "107" in ems_rows
    assert ems_rows["107"][3] == "I32"  # Reactive power setpoint INT32


def test_x7_parse_workbook_direct():
    """Verify parsing function returns non-empty structured data for each subsystem."""
    src = (
        WORKSPACE_DIR
        / "sources"
        / "manuals"
        / "X7"
        / "X7-ModbusTcp-PORT502-20261105-V1.0.6-Beta.xlsx"
    )
    parsed = parse_x7_workbook(src)
    assert len(parsed["ems"]) == 164
    assert len(parsed["bms_summary"]) == 24
    assert len(parsed["bms_cells"]) == 368
    assert len(parsed["meters"]) == 66
    assert len(parsed["pcs"]) == 127
    assert len(parsed["cabinet"]) == 45
    assert len(parsed["all_in_one"]) == 794
