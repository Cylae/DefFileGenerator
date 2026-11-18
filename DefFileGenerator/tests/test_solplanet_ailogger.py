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
    """Verify that COM_Solplanet_Ai-Logger.csv and INV_Solplanet_Ai-Logger.csv satisfy all 5 Apollo pillars."""
    for filename in ("COM_Solplanet_Ai-Logger.csv", "INV_Solplanet_Ai-Logger.csv"):
        prod_file = Path("production") / filename
        assert prod_file.is_file(), f"Missing production file: {prod_file}"

        is_valid, issues = verify_file_pillars(prod_file)
        assert is_valid, f"Verification failed for {filename} with issues: {issues}"
        assert len(issues) == 0


def test_solplanet_ailogger_copies():
    """Verify that all definition repositories contain identical certified copies."""
    for filename in ("COM_Solplanet_Ai-Logger.csv", "INV_Solplanet_Ai-Logger.csv"):
        prod_file = Path("production") / filename
        content = prod_file.read_bytes()

        for folder in ("Definitions", "tagged_definitions", "generated_definitions"):
            target = Path(folder) / filename
            assert target.is_file(), f"Missing copy in {folder}"
            assert target.read_bytes() == content, f"Mismatch between production and {folder}"


def test_solplanet_ailogger_gateway_structure():
    """Verify the internal structure, header, register count, and key telemetry points of the Gateway."""
    prod_file = Path("production/COM_Solplanet_Ai-Logger.csv")
    with open(prod_file, encoding="utf-8") as f:
        reader = list(csv.reader(f, delimiter=";"))

    # Header check
    header = reader[0]
    assert header[0] in ("Modbus", "modbusTCP")
    assert header[1] == "Logger"
    assert header[2] == "Solplanet"
    assert header[3] == "Ai-Logger"

    rows = reader[1:]
    assert len(rows) == 161, f"Expected 161 registers, got {len(rows)}"

    reg_map = {r[2]: r for r in rows}

    # 1. Weather Station (FC 4)
    assert "986" in reg_map
    assert reg_map["986"][1] == "4"
    assert reg_map["986"][3] == "U16"
    assert reg_map["986"][6] == "WIND_SPEED"
    assert reg_map["986"][9] == "m/s"

    assert "987" in reg_map
    assert reg_map["987"][6] == "IRR"

    assert "993" in reg_map
    assert reg_map["993"][6] == "WIND_SPEED_2"
    assert "994" in reg_map
    assert reg_map["994"][6] == "IRR_2"

    # 2. Inverter Telemetry (FC 4)
    assert "1002_32" in reg_map
    assert reg_map["1002_32"][3] == "STRING"
    assert reg_map["1002_32"][6] == "SERIAL"

    assert "1018_16" in reg_map
    assert reg_map["1018_16"][3] == "STRING"
    assert reg_map["1018_16"][6] == "MODEL"

    assert "1027" in reg_map
    assert reg_map["1027"][3] == "U32"
    assert reg_map["1027"][6] == "P_NOM"

    assert "1056_16" in reg_map
    assert reg_map["1056_16"][6] == "MANUFACTURER"

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

    # 6. Direct 4X Inverter Commands
    assert "201" in reg_map
    assert reg_map["201"][6] == "CMD_REMOTE_SWITCH"
    assert "2101" in reg_map
    assert reg_map["2101"][6] == "CMD_MODBUS_ADDR_SET"


def test_solplanet_ailogger_inverter_structure():
    """Verify the internal structure of INV_Solplanet_Ai-Logger.csv."""
    prod_file = Path("production/INV_Solplanet_Ai-Logger.csv")
    with open(prod_file, encoding="utf-8") as f:
        reader = list(csv.reader(f, delimiter=";"))

    header = reader[0]
    assert header[0] in ("Modbus", "modbusTCP")
    assert header[1] == "Inverter"
    assert header[2] == "Solplanet"
    assert header[3] == "Ai-Logger"

    rows = reader[1:]
    assert len(rows) == 106, f"Expected 106 registers, got {len(rows)}"


def test_knowledge_base_contains_ailogger_variables():
    """Verify that dataset_knowledge.json has learned the Ai-Logger variable mappings."""
    kb = load_dataset_knowledge()
    assert kb is not None

    # Check key mappings learned from the Ai-Logger
    assert kb.get_learned_tag("Wind Speed Sensor 1") == "WIND_SPEED"
    assert kb.get_learned_tag("Irradiance Sensor 1") == "IRR"
    assert kb.get_learned_tag("Wind Speed Sensor 2") == "WIND_SPEED_2"
    assert kb.get_learned_tag("Total Inverters Active Power") == "P_AC_TOTAL"
    assert kb.get_learned_tag("Meter Total Active Power") == "P_AC_METER"
    assert (
        kb.get_learned_tag("Active Power Control Command (-100 to 100% Pn)")
        == "CMD_P_AC_SET_PERCENT"
    )


def test_solplanet_field_telemetry_profiles():
    """Verify that real-world SunPM field files for Addr 3..9 match the Method 1 formula."""
    telemetry_dir = Path("sources/telemetry/Solplanet")
    assert telemetry_dir.is_dir(), f"Missing telemetry dir: {telemetry_dir}"

    for addr in range(3, 10):
        field_file = telemetry_dir / f"solplanet_V1_Addr{addr}_PM.csv"
        assert field_file.is_file(), f"Missing field file: {field_file}"

        with open(field_file, encoding="utf-8") as f:
            lines = [line.strip() for line in f if line.strip()]

        # Header check
        header = lines[0].split(";")
        assert header[0] == "modbusTCP"
        assert header[1] == "Inverter"
        assert header[2] == "Solplanet"
        assert header[3] == f"Addr{addr}"

        # 2 registers expected
        assert len(lines) == 3, f"Expected 3 lines (1 header + 2 regs) in {field_file.name}"

        r1 = lines[1].split(";")
        r2 = lines[2].split(";")

        # Method 1 address calculation verification
        expected_addr_reg = str(1001 + (addr - 3) * 390)
        expected_pwr_reg = str(1370 + (addr - 3) * 390)

        # Register 1: Modbus address
        assert r1[0] == "1"
        assert r1[1] == "3"  # Holding register FC 3
        assert r1[2] == expected_addr_reg
        assert r1[3] == "U16"
        assert r1[5] == "Modbus address"
        assert float(r1[7]) == 1.0
        assert r1[10] == "4"  # Action code 4

        # Register 2: Active Power
        assert r2[0] == "2"
        assert r2[1] == "3"  # Holding register FC 3
        assert r2[2] == expected_pwr_reg
        assert r2[3] == "I32"
        assert r2[5] == "Active Power"
        assert float(r2[7]) == 1.0
        assert r2[9] == "W"
        assert r2[10] == "2"  # Action code 2 (Instantaneous telemetry)
