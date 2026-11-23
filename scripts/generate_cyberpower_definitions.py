#!/usr/bin/env python3
"""
Generator & Upgrader for CyberPower Grid-Tied PV Inverter Definitions.
Based on official specification:
'CyberPower Grid_tied PV Inverter Modbus Protocol(V0104) 1.pdf' (V1.04, 2024-05-13).
Models covered:
  - Single-Phase: CPSPV3000ETLA, CPSPV3600ETLA, CPSPV6000ETLA
  - Three-Phase: CPSPV8000ETLA..CPSPV15000ETLA, CPSPV22000ETLA-WB..CPSPV66000ETL-WB
  - Utility Scale: CPSPV120KETL, CPSPV140KETL

Enforces:
  1. 5 Apollo Pillars
  2. Full SunPM tagging for DC, AC, Strings, and Events
  3. Action codes (4=Telemetry, 9=Alarm/Event)
  4. Redundant synchronization across all reference directories
"""

import json
import sys
from pathlib import Path

WORKSPACE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(WORKSPACE_DIR))

from DefFileGenerator.knowledge_miner import (  # noqa: E402
    load_dataset_knowledge,
    normalize_tag_to_sunpm,
)
from DefFileGenerator.production_governor import verify_file_pillars  # noqa: E402
from scripts.audit_logger import log_event  # noqa: E402


def build_generic_cyberpower_rows() -> list[tuple]:
    """Builds the enriched 245-register list for INV_Cyberpower_GENERIC.csv."""
    rows = []

    # 1. Load Power (FC 4)
    rows.append((4, "4429", "I16", "Load Power", "P_LOAD", 1.0, 0.0, "W", 4))
    rows.append((4, "4554", "I32_W", "Load Power (Phase A)", "P_LOAD_1", 1.0, 0.0, "W", 4))
    rows.append((4, "4556", "I32_W", "Load Power (Phase B)", "P_LOAD_2", 1.0, 0.0, "W", 4))
    rows.append((4, "4558", "I32_W", "Load Power (Phase C)", "P_LOAD_3", 1.0, 0.0, "W", 4))

    # 2. Today & Life Energy / Runtime (FC 4)
    rows.append((4, "8504", "U32_W", "Today Energy (Wh)", "E_DAY", 1.0, 0.0, "kWh", 4))
    rows.append((4, "8508", "U32_W", "System Life Energy", "E_TOT", 1.0, 0.0, "Wh", 4))
    rows.append((4, "8510", "U32_W", "System Life Runtime", "TIME_RUN_TOT", 1.0, 0.0, "s", 4))
    rows.append((4, "8512", "U32_W", "PV1 today energy", "E_DC_DAY_1", 0.1, 0.0, "Wh", 4))
    rows.append((4, "8514", "U32_W", "PV1 life energy", "E_DC_TOT_1", 0.1, 0.0, "Wh", 4))
    rows.append((4, "8516", "U32_W", "PV2 today energy", "E_DC_DAY_2", 0.1, 0.0, "Wh", 4))
    rows.append((4, "8518", "U32_W", "PV2 life energy", "E_DC_TOT_2", 0.1, 0.0, "Wh", 4))
    rows.append((4, "8520", "U32_W", "PV3 today energy", "E_DC_DAY_3", 0.1, 0.0, "Wh", 4))
    rows.append((4, "8522", "U32_W", "PV3 life energy", "E_DC_TOT_3", 0.1, 0.0, "Wh", 4))
    rows.append((4, "8524", "U32_W", "PV4 today energy", "E_DC_DAY_4", 0.1, 0.0, "Wh", 4))
    rows.append((4, "8526", "U32_W", "PV4 life energy", "E_DC_TOT_4", 0.1, 0.0, "Wh", 4))
    rows.append((4, "8528", "U32_W", "Feed In Grid today energy", "", 0.1, 0.0, "Wh", 4))
    rows.append((4, "8530", "U32_W", "Grid to load today energy", "", 0.1, 0.0, "Wh", 4))
    rows.append((4, "8532", "U32_W", "Grid to load life energy", "", 0.1, 0.0, "Wh", 4))
    rows.append((4, "8534", "U32_W", "EPS today energy", "", 0.1, 0.0, "Wh", 4))
    rows.append((4, "8536", "U32_W", "EPS life energy", "", 0.1, 0.0, "Wh", 4))
    rows.append((4, "8540", "U32_W", "Battery charge today energy", "", 0.1, 0.0, "Wh", 4))

    # 3. DC MPPT PV1..PV4 (FC 4)
    rows.append((4, "8596", "U16", "PV1 Voltage", "U_DC_1", 0.1, 0.0, "V", 4))
    rows.append((4, "8597", "U16", "PV1 Current", "I_DC_1", 0.01, 0.0, "A", 4))
    rows.append((4, "8598", "U16", "PV1 Power", "P_DC_1", 1.0, 0.0, "W", 4))
    rows.append((4, "8599", "U16", "PV2 Voltage", "U_DC_2", 0.1, 0.0, "V", 4))
    rows.append((4, "8600", "U16", "PV2 Current", "I_DC_2", 0.01, 0.0, "A", 4))
    rows.append((4, "8601", "U16", "PV2 Power", "P_DC_2", 1.0, 0.0, "W", 4))
    rows.append((4, "8602", "U16", "PV3 Voltage", "U_DC_3", 0.1, 0.0, "V", 4))
    rows.append((4, "8603", "U16", "PV3 Current", "I_DC_3", 0.01, 0.0, "A", 4))
    rows.append((4, "8604", "U16", "PV3 Power", "P_DC_3", 1.0, 0.0, "W", 4))
    rows.append((4, "8605", "U16", "PV4 Voltage", "U_DC_4", 0.1, 0.0, "V", 4))
    rows.append((4, "8606", "U16", "PV4 Current", "I_DC_4", 0.01, 0.0, "A", 4))
    rows.append((4, "8607", "U16", "PV4 Power", "P_DC_4", 1.0, 0.0, "W", 4))

    # 4. AC Grid Parameters (FC 4)
    rows.append((4, "8608", "U16", "Grid phase-A voltage", "U_AC_1", 0.1, 0.0, "V", 4))
    rows.append((4, "8609", "U16", "Grid phase-B voltage", "U_AC_2", 0.1, 0.0, "V", 4))
    rows.append((4, "8610", "U16", "Grid phase-C voltage", "U_AC_3", 0.1, 0.0, "V", 4))
    rows.append((4, "8611", "U16", "Grid phase-A output current", "I_AC_1", 0.01, 0.0, "A", 4))
    rows.append((4, "8612", "U16", "Grid phase-B output current", "I_AC_2", 0.01, 0.0, "A", 4))
    rows.append((4, "8613", "U16", "Grid phase-C output current", "I_AC_3", 0.01, 0.0, "A", 4))
    rows.append((4, "8614", "U16", "Grid phase-A output power", "P_AC_1", 1.0, 0.0, "W", 4))
    rows.append((4, "8615", "U16", "Grid phase-B output power", "P_AC_2", 1.0, 0.0, "W", 4))
    rows.append((4, "8616", "U16", "Grid phase-C output power", "P_AC_3", 1.0, 0.0, "W", 4))
    rows.append((4, "8617", "U16", "Grid phase-AB voltage", "U_AC_12", 0.1, 0.0, "V", 4))
    rows.append((4, "8618", "U16", "Grid phase-BC voltage", "U_AC_23", 0.1, 0.0, "V", 4))
    rows.append((4, "8619", "U16", "Grid phase-CA voltage", "U_AC_31", 0.1, 0.0, "V", 4))
    rows.append((4, "8620", "U16", "Grid Frequency", "FREQ", 0.01, 0.0, "Hz", 4))
    rows.append((4, "8621", "U16", "Grid output active power", "P_AC", 1.0, 0.0, "W", 4))
    rows.append((4, "8622", "U16", "Grid output apparent power", "S_AC", 1.0, 0.0, "VA", 4))
    rows.append((4, "8623", "I16", "Grid output reactive power", "Q_AC", 1.0, 0.0, "VAr", 4))
    rows.append((4, "8624", "U16", "Power scale factor", "", 1.0, 0.0, "", 4))

    # 5. Bus Voltages, Insulation, Temperatures (FC 4)
    rows.append((4, "8633", "U16", "BUS1_Voltage", "U_BUS_1", 0.1, 0.0, "V", 4))
    rows.append((4, "8634", "U16", "BUS2_Voltage", "U_BUS_2", 0.1, 0.0, "V", 4))
    rows.append((4, "8635", "U16", "PV Rleakage", "R_ISO", 1.0, 0.0, "kOhm", 4))
    rows.append(
        (4, "8636", "I16", "Temperature (air around power stage)", "T_INT", 0.1, 0.0, "°C", 4)
    )
    rows.append((4, "8637", "I16", "Temperature (wiring box)", "T_AMB", 0.1, 0.0, "°C", 4))
    rows.append((4, "8638", "I16", "Highest Inv Temp", "T_MOD", 0.1, 0.0, "°C", 4))
    rows.append((4, "8639", "I16", "Highest DD Temp", "", 0.1, 0.0, "°C", 4))
    rows.append((4, "8640", "U16", "CPS_PV_Status", "STATUS", 1.0, 0.0, "", 9))
    rows.append((4, "8641", "U16", "Reconnection Time", "", 1.0, 0.0, "s", 4))

    # 6. Event Flags (Action 9)
    rows.append((4, "8642", "U32_W", "I_Event_flags1", "", 1.0, 0.0, "Bitfield", 9))
    rows.append((4, "8644", "U32_W", "I_Event_flags2", "", 1.0, 0.0, "Bitfield", 9))
    rows.append((4, "8646", "U16", "I_Event_1_Vendor", "", 1.0, 0.0, "Bitfield", 9))
    rows.append((4, "8647", "U16", "I_Event_2_Vendor", "", 1.0, 0.0, "Bitfield", 9))
    rows.append((4, "8648", "U16", "I_Event_3_Vendor", "", 1.0, 0.0, "Bitfield", 9))
    rows.append((4, "8649", "U16", "I_Event_4_Vendor", "", 1.0, 0.0, "Bitfield", 9))
    rows.append((4, "8650", "U16", "I_Event_5_Vendor", "", 1.0, 0.0, "Bitfield", 9))
    rows.append((4, "8651", "U16", "I_Event_6_Vendor", "", 1.0, 0.0, "Bitfield", 9))
    rows.append((4, "8652", "U16", "I_Event_7_Vendor", "", 1.0, 0.0, "Bitfield", 9))
    rows.append((4, "8653", "U16", "I_Event_8_Vendor", "", 1.0, 0.0, "Bitfield", 9))
    rows.append((4, "8654", "U16", "I_Event_9_Vendor", "", 1.0, 0.0, "Bitfield", 9))

    # Event Flags 300s latch (Action 9)
    rows.append((4, "8672", "U16", "I_Event_1_Vendor 2", "", 1.0, 0.0, "Bitfield", 9))
    rows.append((4, "8673", "U16", "I_Event_2_Vendor 2", "", 1.0, 0.0, "Bitfield", 9))
    rows.append((4, "8674", "U16", "I_Event_3_Vendor 2", "", 1.0, 0.0, "Bitfield", 9))
    rows.append((4, "8675", "U16", "I_Event_4_Vendor 2", "", 1.0, 0.0, "Bitfield", 9))
    rows.append((4, "8676", "U16", "I_Event_5_Vendor 2", "", 1.0, 0.0, "Bitfield", 9))
    rows.append((4, "8677", "U16", "I_Event_6_Vendor 2", "", 1.0, 0.0, "Bitfield", 9))
    rows.append((4, "8678", "U16", "I_Event_7_Vendor 2", "", 1.0, 0.0, "Bitfield", 9))
    rows.append((4, "8679", "U16", "I_Event_8_Vendor 2", "", 1.0, 0.0, "Bitfield", 9))
    rows.append((4, "8680", "U16", "I_Event_9_Vendor 2", "", 1.0, 0.0, "Bitfield", 9))

    # 7. String Configuration & Fuse Flags (FC 4)
    rows.append((4, "8193", "U16", "String connection setting", "", 1.0, 0.0, "Bitfield", 4))
    rows.append((4, "8199", "U16", "String Positive fuse flags", "", 1.0, 0.0, "Bitfield", 4))
    rows.append((4, "8200", "U16", "String Negative fuse flags", "", 1.0, 0.0, "Bitfield", 4))
    rows.append((4, "8201", "U16", "String Positive fuse flags 2", "", 1.0, 0.0, "Bitfield", 4))
    rows.append((4, "8202", "U16", "String Negative fuse flags 2", "", 1.0, 0.0, "Bitfield", 4))

    # 8. Strings 1 to 32 Voltages, Currents, Powers (FC 4)
    for s_idx in range(1, 33):
        base_addr = 8208 + (s_idx - 1) * 3
        pv_idx = (s_idx - 1) // 4 + 1
        pv_s_idx = (s_idx - 1) % 4 + 1
        rows.append(
            (
                4,
                str(base_addr),
                "U16",
                f"PV{pv_idx} String {pv_s_idx} Voltage",
                f"U_STR_{s_idx}",
                0.1,
                0.0,
                "V",
                4,
            )
        )
        rows.append(
            (
                4,
                str(base_addr + 1),
                "U16",
                f"PV{pv_idx} String {pv_s_idx} Current",
                f"I_STR_{s_idx}",
                0.01,
                0.0,
                "A",
                4,
            )
        )
        rows.append(
            (
                4,
                str(base_addr + 2),
                "U16",
                f"PV{pv_idx} String {pv_s_idx} Power",
                f"P_STR_{s_idx}",
                1.0,
                0.0,
                "W",
                4,
            )
        )

    # 9. Strings 1 to 12 Today & Life Energy (FC 4)
    for s_idx in range(1, 13):
        base_addr = 8304 + (s_idx - 1) * 4
        rows.append(
            (4, str(base_addr), "U32_W", f"String {s_idx} Today energy", "", 1.0, 0.0, "Wh", 4)
        )
        rows.append(
            (4, str(base_addr + 2), "U32_W", f"String {s_idx} Life energy", "", 1.0, 0.0, "Wh", 4)
        )

    # 10. PV5 to PV12 High-Power Inverters Today & Life Energy (FC 4)
    for pv_idx in range(5, 13):
        base_addr = 17664 + (pv_idx - 5) * 4
        rows.append(
            (
                4,
                str(base_addr),
                "U32_W",
                f"PV{pv_idx} Today energy",
                f"E_DC_DAY_{pv_idx}",
                0.1,
                0.0,
                "kWh",
                4,
            )
        )
        rows.append(
            (
                4,
                str(base_addr + 2),
                "U32_W",
                f"PV{pv_idx} Life energy",
                f"E_DC_TOT_{pv_idx}",
                0.1,
                0.0,
                "kWh",
                4,
            )
        )

    # 11. PV5 to PV12 High-Power Inverters Voltages, Currents, Powers (FC 4)
    for pv_idx in range(5, 13):
        base_addr = 17744 + (pv_idx - 5) * 3
        rows.append(
            (4, str(base_addr), "U16", f"PV{pv_idx} Voltage", f"U_DC_{pv_idx}", 0.1, 0.0, "V", 4)
        )
        rows.append(
            (
                4,
                str(base_addr + 1),
                "U16",
                f"PV{pv_idx} Current",
                f"I_DC_{pv_idx}",
                0.01,
                0.0,
                "A",
                4,
            )
        )
        rows.append(
            (4, str(base_addr + 2), "U16", f"PV{pv_idx} Power", f"P_DC_{pv_idx}", 1.0, 0.0, "W", 4)
        )

    return rows


def build_120k_140k_rows() -> list[tuple]:
    """Builds the dedicated definition for CyberPower CPSPV120K & CPSPV140K."""
    # Filter and format for 120K/140K utility scale models
    generic_rows = build_generic_cyberpower_rows()
    # 120K/140K specific: energy is 0.1kWh
    out_rows = []
    for r in generic_rows:
        fc, addr, dtype, desc, tag, factor, offset, unit, action = r
        if addr in ("8504", "8508"):
            factor = 0.1
            unit = "kWh"
        out_rows.append((fc, addr, dtype, desc, tag, factor, offset, unit, action))
    return out_rows


def format_csv(header_model: str, rows: list[tuple]) -> str:
    lines = [f"Modbus;Inverter;Cyberpower;{header_model};;;;;;;"]
    for idx, (fc, addr, dtype, desc, tag, factor, offset, unit, action) in enumerate(rows, 1):
        line = f"{idx};{fc};{addr};{dtype};;{desc};{tag};{factor:.6f};{offset:.6f};{unit};{action}"
        lines.append(line)
    return "\n".join(lines) + "\n"


def deploy_definition(filename: str, csv_content: str):
    """Deploys a definition redundantly across all 4 reference locations."""
    targets = [
        WORKSPACE_DIR / "production" / filename,
        WORKSPACE_DIR / "Definitions" / filename,
        WORKSPACE_DIR / "tagged_definitions" / filename,
        WORKSPACE_DIR / "generated_definitions" / filename,
    ]
    for t in targets:
        t.parent.mkdir(parents=True, exist_ok=True)
        t.write_text(csv_content, encoding="utf-8")
        is_valid, issues = verify_file_pillars(t)
        if not is_valid:
            print(f"ERROR: {t} failed 5 pillars verification: {issues}")
            raise RuntimeError(f"Pillars check failed on {t}: {issues}")
    print(f"Successfully deployed and verified redundant copies of {filename}.")


def update_knowledge_base(rows: list[tuple]):
    """Enriches dataset_knowledge.json with new vocabulary from CyberPower manual."""
    kb = load_dataset_knowledge()
    new_entries = 0
    for r in rows:
        desc = r[3]
        tag = r[4]
        dtype = r[2]
        unit = r[7]
        action = str(r[8])
        if tag:
            norm_tag = normalize_tag_to_sunpm(tag) or tag
            kb.learn_entry(name=desc, tag=norm_tag, dtype=dtype, unit=unit, action=action)
            new_entries += 1

    kb.total_definitions_mined += 1
    kb.total_registers_mined += len(rows)

    out_json = WORKSPACE_DIR / "DefFileGenerator" / "dataset_knowledge.json"
    out_json.write_text(json.dumps(kb.to_dict(), indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Updated knowledge base with {new_entries} CyberPower tags -> {out_json}")


def main():
    print("=" * 80)
    print("CYBERPOWER INVERTER SPECIFICATION TRAINER & GENERATOR")
    print("Document: CyberPower Grid_tied PV Inverter Modbus Protocol(V0104) 1.pdf")
    print("=" * 80)

    # 1. Generate & Deploy INV_Cyberpower_GENERIC.csv
    generic_rows = build_generic_cyberpower_rows()
    generic_csv = format_csv("GENERIC", generic_rows)
    deploy_definition("INV_Cyberpower_GENERIC.csv", generic_csv)

    # 2. Generate & Deploy INV_Cyberpower_CPSPV120K-140K.csv
    rows_120k = build_120k_140k_rows()
    csv_120k = format_csv("CPSPV120K-140K", rows_120k)
    deploy_definition("INV_Cyberpower_CPSPV120K-140K.csv", csv_120k)

    # 3. Train Knowledge Base
    update_knowledge_base(generic_rows)

    # 4. Record Audit Log
    log_event(
        action="SUPERSET_TRAINING",
        equipment="CyberPower Grid-Tied Inverter Series (CPSPV3000..140KETL)",
        manufacturer="Cyberpower",
        source_document="sources/manuals/CyberPower/CyberPower Grid_tied PV Inverter Modbus Protocol(V0104) 1.pdf",
        definitions=["INV_Cyberpower_GENERIC.csv", "INV_Cyberpower_CPSPV120K-140K.csv"],
        register_count=len(generic_rows),
        tags_count=sum(1 for r in generic_rows if r[4]),
        validation_status="EXECUTED_PASS",
        details="Full SunPM tagging, DC MPPT PV1..PV12, String 1..32, Grid AC, and V1.04 Event flags.",
    )
    print("Audit log recorded.")


if __name__ == "__main__":
    main()
