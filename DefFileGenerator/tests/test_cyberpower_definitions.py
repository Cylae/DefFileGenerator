#!/usr/bin/env python3
"""
Unit and Verification Tests for CyberPower Grid-Tied PV Inverter Definitions (Protocol V1.04).
Tests 5-pillar compliance, register completeness, SunPM tagging, and redundancy.
"""

import csv
from pathlib import Path

from DefFileGenerator.knowledge_miner import load_dataset_knowledge
from DefFileGenerator.production_governor import verify_file_pillars


def test_cyberpower_generic_pillars():
    """Verify that INV_Cyberpower_GENERIC.csv complies with 5 Apollo pillars."""
    prod_file = Path("production/INV_Cyberpower_GENERIC.csv")
    assert prod_file.is_file(), f"Missing {prod_file}"

    is_valid, issues = verify_file_pillars(prod_file)
    assert is_valid, f"Verification failed on {prod_file}: {issues}"


def test_cyberpower_120k_140k_pillars():
    """Verify that INV_Cyberpower_CPSPV120K-140K.csv complies with 5 Apollo pillars."""
    prod_file = Path("production/INV_Cyberpower_CPSPV120K-140K.csv")
    assert prod_file.is_file(), f"Missing {prod_file}"

    is_valid, issues = verify_file_pillars(prod_file)
    assert is_valid, f"Verification failed on {prod_file}: {issues}"


def test_cyberpower_redundancy():
    """Verify that CyberPower definitions are byte-identical across all 4 directories."""
    filenames = ["INV_Cyberpower_GENERIC.csv", "INV_Cyberpower_CPSPV120K-140K.csv"]
    folders = ["Definitions", "tagged_definitions", "generated_definitions"]

    for fname in filenames:
        prod_bytes = (Path("production") / fname).read_bytes()
        for fld in folders:
            target = Path(fld) / fname
            assert target.is_file(), f"Missing {target}"
            assert target.read_bytes() == prod_bytes, f"Divergence in {target}"


def test_cyberpower_telemetry_structure():
    """Verify detailed register mapping, types, units, and SunPM telemetry tags."""
    prod_file = Path("production/INV_Cyberpower_GENERIC.csv")
    with open(prod_file, encoding="utf-8") as f:
        reader = list(csv.reader(f, delimiter=";"))

    rows = reader[1:]
    assert len(rows) == 244, f"Expected 244 rows, got {len(rows)}"

    reg_map = {r[2]: r for r in rows}

    # 1. DC MPPT PV1..PV12
    for idx in range(1, 13):
        assert any(r[6] == f"U_DC_{idx}" for r in rows), f"Missing U_DC_{idx}"
        assert any(r[6] == f"I_DC_{idx}" for r in rows), f"Missing I_DC_{idx}"
        assert any(r[6] == f"P_DC_{idx}" for r in rows), f"Missing P_DC_{idx}"

    # 2. AC Grid
    assert reg_map["8608"][6] == "U_AC_1"
    assert reg_map["8609"][6] == "U_AC_2"
    assert reg_map["8610"][6] == "U_AC_3"
    assert reg_map["8611"][6] == "I_AC_1"
    assert reg_map["8612"][6] == "I_AC_2"
    assert reg_map["8613"][6] == "I_AC_3"
    assert reg_map["8614"][6] == "P_AC_1"
    assert reg_map["8614"][9] == "W"  # Fixed unit W (was A)
    assert reg_map["8620"][6] == "FREQ"
    assert reg_map["8621"][6] == "P_AC"
    assert reg_map["8622"][6] == "S_AC"
    assert reg_map["8623"][6] == "Q_AC"

    # 3. Strings 1..32
    for s_idx in range(1, 33):
        assert any(r[6] == f"I_STR_{s_idx}" for r in rows), f"Missing I_STR_{s_idx}"

    # 4. Action Code 9 for Events
    for ev_addr in ("8640", "8642", "8644", "8646", "8647", "8648", "8649", "8650", "8651", "8652", "8653", "8654"):
        assert reg_map[ev_addr][10] == "9", f"Event register {ev_addr} should have Action 9"


def test_cyberpower_knowledge_learning():
    """Verify that dataset_knowledge.json has learned the CyberPower vocabulary."""
    kb = load_dataset_knowledge()
    assert kb is not None
    assert kb.get_learned_tag("Grid output active power") == "P_AC"
    assert kb.get_learned_tag("Grid output apparent power") == "S_AC"
    assert kb.get_learned_tag("Grid output reactive power") == "Q_AC"
    assert kb.get_learned_tag("PV1 Voltage") == "U_DC_1"
    assert kb.get_learned_tag("PV1 Current") == "I_DC_1"
