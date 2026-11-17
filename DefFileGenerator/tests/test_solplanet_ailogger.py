#!/usr/bin/env python3
"""
Unit and Contract Verification Tests for Solplanet Ai-Logger Definition.
Validates definition compliance with the 5 Apollo Pillars and dataset knowledge base integration.
"""

import csv
from pathlib import Path

from DefFileGenerator.knowledge_miner import load_dataset_knowledge
from DefFileGenerator.production_governor import verify_file_pillars


def test_solplanet_ailogger_pillars():
    """Verify that COM_Solplanet_Ai-Logger.csv satisfies all 5 Apollo pillars."""
    prod_file = Path("production/COM_Solplanet_Ai-Logger.csv")
    assert prod_file.is_file(), f"Missing production file: {prod_file}"

    is_valid, issues = verify_file_pillars(prod_file)
    assert is_valid, f"Verification failed with issues: {issues}"
    assert len(issues) == 0


def test_solplanet_ailogger_copies():
    """Verify that all definition repositories contain identical certified copies."""
    prod_file = Path("production/COM_Solplanet_Ai-Logger.csv")
    content = prod_file.read_bytes()

    for folder in ("Definitions", "tagged_definitions", "generated_definitions"):
        target = Path(folder) / "COM_Solplanet_Ai-Logger.csv"
        assert target.is_file(), f"Missing copy in {folder}"
        assert target.read_bytes() == content, f"Mismatch between production and {folder}"


def test_solplanet_ailogger_structure():
    """Verify the internal structure, header, register count, and key telemetry points."""
    prod_file = Path("production/COM_Solplanet_Ai-Logger.csv")
    with open(prod_file, encoding="utf-8") as f:
        reader = list(csv.reader(f, delimiter=";"))

    # Header check
    header = reader[0]
    assert header[0] == "Modbus"
    assert header[1] == "Logger"
    assert header[2] == "Solplanet"
    assert header[3] == "Ai-Logger"

    rows = reader[1:]
    assert len(rows) == 159, f"Expected 159 registers, got {len(rows)}"

    reg_map = {r[2]: r for r in rows}

    # 1. Weather Station (FC 4)
    assert "986" in reg_map
    assert reg_map["986"][1] == "4"
    assert reg_map["986"][3] == "U16"
    assert reg_map["986"][6] == "WIND_SPEED"
    assert reg_map["986"][9] == "m/s"

    assert "987" in reg_map
    assert reg_map["987"][6] == "IRR"

    # 2. Inverter Telemetry (FC 4)
    assert "1002_16" in reg_map
    assert reg_map["1002_16"][3] == "STRING"
    assert reg_map["1002_16"][6] == "INV_SN"

    assert "1027" in reg_map
    assert reg_map["1027"][3] == "U32"
    assert reg_map["1027"][6] == "P_NOM"

    assert "1308" in reg_map
    assert reg_map["1308"][6] == "STATUS"
    assert reg_map["1308"][10] == "9"  # Alarm/Status

    assert "1370" in reg_map
    assert reg_map["1370"][3] == "I32"
    assert reg_map["1370"][6] == "P_AC"

    # 3. Total Inverter Power (FC 4)
    assert "36100" in reg_map
    assert reg_map["36100"][3] == "I32"
    assert reg_map["36100"][6] == "P_AC_TOTAL"

    # 4. SDM630 Meter (FC 4, F32)
    assert "36104" in reg_map
    assert reg_map["36104"][3] == "F32"
    assert reg_map["36104"][6] == "U_AC_METER_1"

    assert "36156" in reg_map
    assert reg_map["36156"][3] == "F32"
    assert reg_map["36156"][6] == "P_AC_METER"

    assert "36176" in reg_map
    assert reg_map["36176"][3] == "F32"
    assert reg_map["36176"][6] == "E_IMP"

    # 5. Remote Inverter Control (FC 3, Action 1)
    assert "60000" in reg_map
    assert reg_map["60000"][1] == "3"  # Holding register
    assert reg_map["60000"][3] == "I16"
    assert reg_map["60000"][6] == "CMD_P_AC_SET_PERCENT"
    assert reg_map["60000"][10] == "1"  # Command/RW

    assert "60002" in reg_map
    assert reg_map["60002"][1] == "3"
    assert reg_map["60002"][6] == "CMD_ON_OFF"
    assert reg_map["60002"][10] == "1"


def test_knowledge_base_contains_ailogger_variables():
    """Verify that dataset_knowledge.json has learned the Ai-Logger variable mappings."""
    kb = load_dataset_knowledge()
    assert kb is not None

    # Check key mappings learned from the Ai-Logger
    assert kb.get_learned_tag("Wind Speed Sensor 1") == "WIND_SPEED"
    assert kb.get_learned_tag("Irradiance Sensor 1") == "IRR"
    assert kb.get_learned_tag("Total Inverters Active Power") == "P_AC_TOTAL"
    assert kb.get_learned_tag("Meter Total Active Power") == "P_AC_METER"
    assert kb.get_learned_tag("Active Power Control Command (-100 to 100% Pn)") == "CMD_P_AC_SET_PERCENT"
