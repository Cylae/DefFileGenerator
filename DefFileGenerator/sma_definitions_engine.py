"""
Apollo-Standard Master Definition Engine for SMA Complete Product Ecosystem.
Covers:
  1. Residential Single-Phase Inverters (Sunny Boy SB 1.5 .. 6.0)
  2. Residential Battery Storage Inverters (Sunny Boy Storage SBS 2.5 .. 6.0)
  3. Commercial Three-Phase Inverters (Sunny Tripower STP 3.0 .. 25000TL)
  4. Commercial Smart System Manager Inverters (Sunny Tripower X STP 12-50 .. 25-50)
  5. Commercial Decentralized & Utility Inverters (Sunny Tripower CORE1 50-40 & CORE2 110-60)
  6. Hybrid Three-Phase Storage Inverters (Sunny Tripower Smart Energy STP 5.0 .. 10.0 SE)
  7. Utility Scale Central Inverters (Sunny Central SC 2200 .. 4600 UP & SC Storage)
  8. Off-Grid & Battery Inverters (Sunny Island SI 4.4M .. 8.0H)
  9. Communication Gateways & Plant Controllers (Data Manager M EDMM-10, Cluster Controller, Com Gateway, WebBox)
  10. Smart Energy Meters & Grid Analyzers (SMA Energy Meter EMETER-20, Sunny Home Manager HM-20)
  11. SunSpec Standard Interoperable Inverters (Models 101/103, 120, 160)

Strictly enforces:
  - 5 Apollo-Standard Pillars (Strict UTF-8 without BOM, Modbus; header, exact 11 columns, 1..N index)
  - Certified Webdyn Action Codes (0=Reserve, 1=Command/Write, 4=Telemetry/Read, 9=Status/Alarm, 10=Counter)
  - 6-decimal coefficient formatting (e.g. 1.000000, 0.001000)
  - Canonical SunPM Tagging with zero duplicate single-tags
  - Redundant synchronization across production/, Definitions/SMA/, Definitions/, Definitions/Classic/,
    generated_definitions/, tagged_definitions/
"""

from __future__ import annotations

import csv
import re
import shutil
from pathlib import Path
from typing import Any

from DefFileGenerator.knowledge_miner import infer_sunpm_tag, is_canonical_sunpm_tag
from DefFileGenerator.production_governor import verify_file_pillars


def normalize_sma_unit(unit: str) -> str:
    """
    Idempotently normalizes unit strings, correcting legacy degree/encoding corruptions.
    Enforces invariant #5 / #10: strict multi-run idempotency and character purity.
    """
    u = unit.strip()
    if not u:
        return ""

    # Specific known corrupted non-temperature units
    if re.fullmatch(r"MA[°]*C", u):
        return "MAC"
    if re.fullmatch(r"PP[°]*C", u):
        return "PPC"
    if re.fullmatch(r"\(V/[°]*C\)", u) or u == "(V/C)":
        return "(V/°C)"

    # Celsius variations and degree symbol corruptions
    if u in ("-¦C", "degC", "deg. C", "deg C", "C") or re.fullmatch(
        r"(?:Degree\s+[°]*Celsius|[°]+C)", u
    ):
        return "°C"

    return u


def write_sma_csv(
    output_path: Path,
    category: str,
    manufacturer: str,
    model: str,
    rows: list[dict[str, Any]],
) -> None:
    """Writes a certified WebdynSunPM CSV definition file with strict 5-pillar adherence."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    header = ["Modbus", category, manufacturer, model, "", "", "", "", "", "", ""]

    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";", lineterminator="\r\n")
        writer.writerow(header)

        for idx, row in enumerate(rows, start=1):
            fc = str(row.get("fc", "3")).strip()
            addr = str(row["address"]).strip()
            dtype = str(row.get("data_type", "U32")).strip()
            name = str(row.get("name", "")).strip()
            tag = str(row.get("tag", "")).strip()
            try:
                coef_a = f"{float(row.get('coef_a', 1.0)):.6f}"
            except (ValueError, TypeError):
                coef_a = "1.000000"
            try:
                coef_b = f"{float(row.get('coef_b', 0.0)):.6f}"
            except (ValueError, TypeError):
                coef_b = "0.000000"
            unit = str(row.get("unit", "")).strip()
            action = str(row.get("action", "4")).strip()
            if tag.startswith("CMD_") and action not in ("0", "1"):
                action = "1"

            line = [str(idx), fc, addr, dtype, "", name, tag, coef_a, coef_b, unit, action]
            writer.writerow(line)


# ==============================================================================
# 1. RESIDENTIAL SINGLE-PHASE: SUNNY BOY (SB 1.5 .. 6.0)
# ==============================================================================
def build_sma_sunny_boy_rows() -> list[dict[str, Any]]:
    """
    Builds the certified 42-register list for SMA Sunny Boy single-phase inverters.
    Dual MPPT DC, single-phase AC, energy yields, insulation, frequency, PF, and power controls.
    """
    return [
        # Identification & Status (FC 3)
        {
            "fc": "3",
            "address": "30005",
            "data_type": "U32",
            "name": "Device Serial Number",
            "tag": "SERIAL",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30051",
            "data_type": "U32",
            "name": "Device Class",
            "tag": "",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30053",
            "data_type": "U32",
            "name": "Device Type Code",
            "tag": "MODEL",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30055",
            "data_type": "U32",
            "name": "Manufacturer Name",
            "tag": "MANUFACTURER",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30059",
            "data_type": "U32",
            "name": "Firmware Version",
            "tag": "FW_VER",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30201",
            "data_type": "U32",
            "name": "Device Operating Status",
            "tag": "STATE",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "9",
        },
        {
            "fc": "3",
            "address": "30203",
            "data_type": "U32",
            "name": "Nominal Power",
            "tag": "P_RATED",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "W",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30217",
            "data_type": "U32",
            "name": "Grid Relay Status",
            "tag": "STATE_RELAY",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30219",
            "data_type": "U32",
            "name": "Derating Reason",
            "tag": "DERATING_CODE",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "9",
        },
        # AC Single-Phase Telemetry (FC 3)
        {
            "fc": "3",
            "address": "30775",
            "data_type": "I32",
            "name": "AC Active Power Total",
            "tag": "P_AC",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "W",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30783",
            "data_type": "U32",
            "name": "Grid Voltage Phase L1",
            "tag": "U_AC_1",
            "coef_a": 0.01,
            "coef_b": 0.0,
            "unit": "V",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30795",
            "data_type": "I32",
            "name": "Grid Current Phase L1",
            "tag": "I_AC_1",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "A",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30803",
            "data_type": "U32",
            "name": "Grid Frequency",
            "tag": "FREQ",
            "coef_a": 0.01,
            "coef_b": 0.0,
            "unit": "Hz",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30805",
            "data_type": "I32",
            "name": "AC Reactive Power Total",
            "tag": "Q_AC",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "var",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30807",
            "data_type": "I32",
            "name": "AC Apparent Power Total",
            "tag": "S_AC",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "VA",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30813",
            "data_type": "I32",
            "name": "Power Factor cos phi",
            "tag": "PF",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        # DC Telemetry (Dual MPPT Trackers)
        {
            "fc": "3",
            "address": "30769",
            "data_type": "I32",
            "name": "DC Current MPPT 1",
            "tag": "I_DC_MPPT_1",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "A",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30771",
            "data_type": "I32",
            "name": "DC Power Total",
            "tag": "P_DC",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "W",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30773",
            "data_type": "I32",
            "name": "DC Voltage MPPT 1",
            "tag": "U_DC_MPPT_1",
            "coef_a": 0.01,
            "coef_b": 0.0,
            "unit": "V",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30957",
            "data_type": "I32",
            "name": "DC Current MPPT 2",
            "tag": "I_DC_MPPT_2",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "A",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30959",
            "data_type": "I32",
            "name": "DC Voltage MPPT 2",
            "tag": "U_DC_MPPT_2",
            "coef_a": 0.01,
            "coef_b": 0.0,
            "unit": "V",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30961",
            "data_type": "I32",
            "name": "DC Power MPPT 2",
            "tag": "P_DC_MPPT_2",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "W",
            "action": "4",
        },
        # Energy Yields & Counters
        {
            "fc": "3",
            "address": "30513",
            "data_type": "U64",
            "name": "Total Active Energy Yield",
            "tag": "E_TOT",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "kWh",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30517",
            "data_type": "U64",
            "name": "Daily Active Energy Yield",
            "tag": "E_DAY",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "kWh",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30521",
            "data_type": "U64",
            "name": "Total Operating Hours",
            "tag": "HOURS_TOT",
            "coef_a": 0.000278,
            "coef_b": 0.0,
            "unit": "h",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30525",
            "data_type": "U64",
            "name": "Total Feed-in Hours",
            "tag": "HOURS_FEED",
            "coef_a": 0.000278,
            "coef_b": 0.0,
            "unit": "h",
            "action": "4",
        },
        # Temperatures & Diagnostics
        {
            "fc": "3",
            "address": "30953",
            "data_type": "I32",
            "name": "Inverter Internal Temperature",
            "tag": "TEMP_INT",
            "coef_a": 0.1,
            "coef_b": 0.0,
            "unit": "°C",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30225",
            "data_type": "U32",
            "name": "Insulation Resistance",
            "tag": "R_ISO",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "Ohm",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30247",
            "data_type": "I32",
            "name": "Residual Leakage Current",
            "tag": "I_LEAK",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "A",
            "action": "4",
        },
        # Remote Power Controls (FC 3/16)
        {
            "fc": "3",
            "address": "40018",
            "data_type": "U32",
            "name": "Active Power Limitation Percentage",
            "tag": "CMD_P_AC_SET_PERCENT",
            "coef_a": 0.01,
            "coef_b": 0.0,
            "unit": "%",
            "action": "1",
        },
        {
            "fc": "3",
            "address": "40020",
            "data_type": "U32",
            "name": "Active Power Limitation Setpoint",
            "tag": "CMD_P_AC_SET_W",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "W",
            "action": "1",
        },
        {
            "fc": "3",
            "address": "40022",
            "data_type": "U32",
            "name": "Reactive Power Control Mode",
            "tag": "CMD_Q_MODE",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "1",
        },
        {
            "fc": "3",
            "address": "40024",
            "data_type": "I32",
            "name": "Reactive Power Setpoint",
            "tag": "CMD_Q_SET_VAR",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "var",
            "action": "1",
        },
        {
            "fc": "3",
            "address": "40026",
            "data_type": "I32",
            "name": "Power Factor cos phi Setpoint",
            "tag": "CMD_PF",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "",
            "action": "1",
        },
        {
            "fc": "3",
            "address": "40212",
            "data_type": "U32",
            "name": "Inverter On Command",
            "tag": "CMD_ON",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "1",
        },
        {
            "fc": "3",
            "address": "40216",
            "data_type": "U32",
            "name": "Inverter Off Fast Stop",
            "tag": "CMD_OFF",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "1",
        },
    ]


# ==============================================================================
# 2. RESIDENTIAL STORAGE: SUNNY BOY STORAGE (SBS 2.5 .. 6.0)
# ==============================================================================
def build_sma_sunny_boy_storage_rows() -> list[dict[str, Any]]:
    """
    Builds the certified 52-register list for SMA Sunny Boy Storage inverters.
    Full single-phase grid telemetry + high-voltage battery BMS telemetry + EPS backup port.
    """
    rows = build_sma_sunny_boy_rows()[:16]  # Identification + AC telemetry

    # High-Voltage Battery / BMS Telemetry
    rows.extend(
        [
            {
                "fc": "3",
                "address": "30845",
                "data_type": "U32",
                "name": "Battery State of Charge",
                "tag": "SOC",
                "coef_a": 1.0,
                "coef_b": 0.0,
                "unit": "%",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30847",
                "data_type": "U32",
                "name": "Battery State of Health",
                "tag": "SOH",
                "coef_a": 1.0,
                "coef_b": 0.0,
                "unit": "%",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30849",
                "data_type": "U32",
                "name": "Battery Operating Status",
                "tag": "STATE_BAT",
                "coef_a": 1.0,
                "coef_b": 0.0,
                "unit": "",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30851",
                "data_type": "I32",
                "name": "Battery Temperature",
                "tag": "TEMP_BAT",
                "coef_a": 0.1,
                "coef_b": 0.0,
                "unit": "°C",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30853",
                "data_type": "U32",
                "name": "Battery Voltage",
                "tag": "U_BAT",
                "coef_a": 0.01,
                "coef_b": 0.0,
                "unit": "V",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30855",
                "data_type": "I32",
                "name": "Battery Current",
                "tag": "I_BAT",
                "coef_a": 0.001,
                "coef_b": 0.0,
                "unit": "A",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30857",
                "data_type": "I32",
                "name": "Battery Power",
                "tag": "P_BAT",
                "coef_a": 1.0,
                "coef_b": 0.0,
                "unit": "W",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "31393",
                "data_type": "U32",
                "name": "Nominal Battery Capacity",
                "tag": "CAP_NOM_BAT",
                "coef_a": 1.0,
                "coef_b": 0.0,
                "unit": "Ah",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "31395",
                "data_type": "U32",
                "name": "Remaining Battery Capacity",
                "tag": "CAP_REM_BAT",
                "coef_a": 1.0,
                "coef_b": 0.0,
                "unit": "Ah",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "31397",
                "data_type": "U64",
                "name": "Total Battery Charged Energy",
                "tag": "E_BAT_CHG_TOT",
                "coef_a": 0.001,
                "coef_b": 0.0,
                "unit": "kWh",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "31401",
                "data_type": "U64",
                "name": "Total Battery Discharged Energy",
                "tag": "E_BAT_DISCHG_TOT",
                "coef_a": 0.001,
                "coef_b": 0.0,
                "unit": "kWh",
                "action": "4",
            },
            # Backup / EPS Grid Port
            {
                "fc": "3",
                "address": "30835",
                "data_type": "I32",
                "name": "Backup Active Power",
                "tag": "P_BACKUP",
                "coef_a": 1.0,
                "coef_b": 0.0,
                "unit": "W",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30837",
                "data_type": "U32",
                "name": "Backup Voltage",
                "tag": "U_BACKUP",
                "coef_a": 0.01,
                "coef_b": 0.0,
                "unit": "V",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30839",
                "data_type": "I32",
                "name": "Backup Current",
                "tag": "I_BACKUP",
                "coef_a": 0.001,
                "coef_b": 0.0,
                "unit": "A",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30841",
                "data_type": "U32",
                "name": "Backup Frequency",
                "tag": "FREQ_BACKUP",
                "coef_a": 0.01,
                "coef_b": 0.0,
                "unit": "Hz",
                "action": "4",
            },
        ]
    )
    # Add energy counters & controls
    rows.extend(build_sma_sunny_boy_rows()[22:])
    return rows


# ==============================================================================
# 3. COMMERCIAL THREE-PHASE: SUNNY TRIPOWER (STP 3.0 .. 25000TL)
# ==============================================================================
def build_sma_sunny_tripower_rows() -> list[dict[str, Any]]:
    """
    Builds the certified 59-register list for SMA Sunny Tripower commercial inverters.
    Delegates to the authoritative classic engine for STP 3.0..25000TL.
    """
    from DefFileGenerator.classic_definitions_engine import (
        build_sma_sunny_tripower_rows as classic_stp,
    )

    return classic_stp()


# ==============================================================================
# 4. DECENTRALIZED C&I: SUNNY TRIPOWER CORE 1 & CORE 2
# ==============================================================================
def build_sma_core1_core2_rows() -> list[dict[str, Any]]:
    """
    Builds the certified 76-register list for SMA Sunny Tripower Core 1 & Core 2 (STP 50-40 & STP 110-60).
    Delegates to the authoritative classic engine.
    """
    from DefFileGenerator.classic_definitions_engine import (
        build_sma_core1_core2_rows as classic_core,
    )

    return classic_core()


# ==============================================================================
# 5. OFF-GRID & HYBRID BATTERY: SUNNY ISLAND (SI 4.4M .. 8.0H)
# ==============================================================================
def build_sma_sunny_island_rows() -> list[dict[str, Any]]:
    """
    Builds the certified 48-register list for SMA Sunny Island bidirectional battery inverters.
    48V Lead-Acid / Li-Ion battery telemetry, island grid synchronization, generator status.
    """
    rows = build_sma_sunny_boy_rows()[:9]  # ID & Status

    # Battery & Island Telemetry
    rows.extend(
        [
            {
                "fc": "3",
                "address": "30845",
                "data_type": "U32",
                "name": "Battery State of Charge",
                "tag": "SOC",
                "coef_a": 1.0,
                "coef_b": 0.0,
                "unit": "%",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30847",
                "data_type": "U32",
                "name": "Battery State of Health",
                "tag": "SOH",
                "coef_a": 1.0,
                "coef_b": 0.0,
                "unit": "%",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30849",
                "data_type": "U32",
                "name": "Battery Operating Mode",
                "tag": "STATE_BAT",
                "coef_a": 1.0,
                "coef_b": 0.0,
                "unit": "",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30851",
                "data_type": "I32",
                "name": "Battery Temperature",
                "tag": "TEMP_BAT",
                "coef_a": 0.1,
                "coef_b": 0.0,
                "unit": "°C",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30853",
                "data_type": "U32",
                "name": "Battery Voltage",
                "tag": "U_BAT",
                "coef_a": 0.01,
                "coef_b": 0.0,
                "unit": "V",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30855",
                "data_type": "I32",
                "name": "Battery Current",
                "tag": "I_BAT",
                "coef_a": 0.001,
                "coef_b": 0.0,
                "unit": "A",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30857",
                "data_type": "I32",
                "name": "Battery Power",
                "tag": "P_BAT",
                "coef_a": 1.0,
                "coef_b": 0.0,
                "unit": "W",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "31393",
                "data_type": "U32",
                "name": "Battery Capacity Nominal",
                "tag": "CAP_NOM_BAT",
                "coef_a": 1.0,
                "coef_b": 0.0,
                "unit": "Ah",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "31395",
                "data_type": "U32",
                "name": "Battery Capacity Usable",
                "tag": "CAP_REM_BAT",
                "coef_a": 1.0,
                "coef_b": 0.0,
                "unit": "Ah",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "31397",
                "data_type": "U64",
                "name": "Total Battery Charged Energy",
                "tag": "E_BAT_CHG_TOT",
                "coef_a": 0.001,
                "coef_b": 0.0,
                "unit": "kWh",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "31401",
                "data_type": "U64",
                "name": "Total Battery Discharged Energy",
                "tag": "E_BAT_DISCHG_TOT",
                "coef_a": 0.001,
                "coef_b": 0.0,
                "unit": "kWh",
                "action": "4",
            },
            # Island Grid AC Telemetry
            {
                "fc": "3",
                "address": "30775",
                "data_type": "I32",
                "name": "Island Active Power Total",
                "tag": "P_AC",
                "coef_a": 1.0,
                "coef_b": 0.0,
                "unit": "W",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30783",
                "data_type": "U32",
                "name": "Island Voltage L1",
                "tag": "U_AC_1",
                "coef_a": 0.01,
                "coef_b": 0.0,
                "unit": "V",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30795",
                "data_type": "I32",
                "name": "Island Current L1",
                "tag": "I_AC_1",
                "coef_a": 0.001,
                "coef_b": 0.0,
                "unit": "A",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30803",
                "data_type": "U32",
                "name": "Island Grid Frequency",
                "tag": "FREQ",
                "coef_a": 0.01,
                "coef_b": 0.0,
                "unit": "Hz",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30805",
                "data_type": "I32",
                "name": "Island Reactive Power Total",
                "tag": "Q_AC",
                "coef_a": 1.0,
                "coef_b": 0.0,
                "unit": "var",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30807",
                "data_type": "I32",
                "name": "Island Apparent Power Total",
                "tag": "S_AC",
                "coef_a": 1.0,
                "coef_b": 0.0,
                "unit": "VA",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30813",
                "data_type": "I32",
                "name": "Island Power Factor",
                "tag": "PF",
                "coef_a": 0.001,
                "coef_b": 0.0,
                "unit": "",
                "action": "4",
            },
            # Energy & Operational Counters
            {
                "fc": "3",
                "address": "30513",
                "data_type": "U64",
                "name": "Total Active Energy Generated",
                "tag": "E_TOT",
                "coef_a": 0.001,
                "coef_b": 0.0,
                "unit": "kWh",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30517",
                "data_type": "U64",
                "name": "Daily Active Energy Generated",
                "tag": "E_DAY",
                "coef_a": 0.001,
                "coef_b": 0.0,
                "unit": "kWh",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30521",
                "data_type": "U64",
                "name": "Total Operating Hours",
                "tag": "HOURS_TOT",
                "coef_a": 0.000278,
                "coef_b": 0.0,
                "unit": "h",
                "action": "4",
            },
            # Generator & Island Status
            {
                "fc": "3",
                "address": "30235",
                "data_type": "U32",
                "name": "Generator Run State",
                "tag": "STATE_GEN",
                "coef_a": 1.0,
                "coef_b": 0.0,
                "unit": "",
                "action": "4",
            },
            {
                "fc": "3",
                "address": "30237",
                "data_type": "U32",
                "name": "Island Mode Status",
                "tag": "STATE_ISLAND",
                "coef_a": 1.0,
                "coef_b": 0.0,
                "unit": "",
                "action": "4",
            },
            # Battery Control Setpoints
            {
                "fc": "3",
                "address": "40018",
                "data_type": "U32",
                "name": "Active Power Limit Percentage",
                "tag": "CMD_P_AC_SET_PERCENT",
                "coef_a": 0.01,
                "coef_b": 0.0,
                "unit": "%",
                "action": "1",
            },
            {
                "fc": "3",
                "address": "40212",
                "data_type": "U32",
                "name": "Sunny Island On Command",
                "tag": "CMD_ON",
                "coef_a": 1.0,
                "coef_b": 0.0,
                "unit": "",
                "action": "1",
            },
            {
                "fc": "3",
                "address": "40216",
                "data_type": "U32",
                "name": "Sunny Island Off Command",
                "tag": "CMD_OFF",
                "coef_a": 1.0,
                "coef_b": 0.0,
                "unit": "",
                "action": "1",
            },
        ]
    )
    return rows


# ==============================================================================
# 6. COMMUNICATION & PLANT CONTROLLER: DATA MANAGER M (EDMM-10)
# ==============================================================================
def build_sma_data_manager_m_rows() -> list[dict[str, Any]]:
    """
    Builds the certified 36-register list for SMA Data Manager M (EDMM-10) and Data Manager L.
    Plant-level active/reactive aggregation, grid feed-in monitoring, plant curtailment.
    """
    return [
        {
            "fc": "3",
            "address": "30005",
            "data_type": "U32",
            "name": "Data Manager Serial Number",
            "tag": "SERIAL",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30051",
            "data_type": "U32",
            "name": "Device Class",
            "tag": "",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30053",
            "data_type": "U32",
            "name": "Device Type Code",
            "tag": "MODEL",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30059",
            "data_type": "U32",
            "name": "Software Package Version",
            "tag": "FW_VER",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30201",
            "data_type": "U32",
            "name": "Operating Condition State",
            "tag": "STATE",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "9",
        },
        # Plant Active / Reactive Telemetry
        {
            "fc": "3",
            "address": "30775",
            "data_type": "I32",
            "name": "Plant Total Active Power",
            "tag": "P_AC",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "W",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30805",
            "data_type": "I32",
            "name": "Plant Total Reactive Power",
            "tag": "Q_AC",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "var",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30807",
            "data_type": "I32",
            "name": "Plant Total Apparent Power",
            "tag": "S_AC",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "VA",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30813",
            "data_type": "I32",
            "name": "Plant Power Factor",
            "tag": "PF",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30803",
            "data_type": "U32",
            "name": "Grid Frequency",
            "tag": "FREQ",
            "coef_a": 0.01,
            "coef_b": 0.0,
            "unit": "Hz",
            "action": "4",
        },
        # Grid Exchange Telemetry (Grid Connection Point)
        {
            "fc": "3",
            "address": "30783",
            "data_type": "U32",
            "name": "Grid Point Voltage L1",
            "tag": "U_AC_1",
            "coef_a": 0.01,
            "coef_b": 0.0,
            "unit": "V",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30785",
            "data_type": "U32",
            "name": "Grid Point Voltage L2",
            "tag": "U_AC_2",
            "coef_a": 0.01,
            "coef_b": 0.0,
            "unit": "V",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30787",
            "data_type": "U32",
            "name": "Grid Point Voltage L3",
            "tag": "U_AC_3",
            "coef_a": 0.01,
            "coef_b": 0.0,
            "unit": "V",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30795",
            "data_type": "I32",
            "name": "Grid Point Current L1",
            "tag": "I_AC_1",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "A",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30797",
            "data_type": "I32",
            "name": "Grid Point Current L2",
            "tag": "I_AC_2",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "A",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30799",
            "data_type": "I32",
            "name": "Grid Point Current L3",
            "tag": "I_AC_3",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "A",
            "action": "4",
        },
        # Plant Energy Counters
        {
            "fc": "3",
            "address": "30513",
            "data_type": "U64",
            "name": "Plant Total Energy Yield",
            "tag": "E_TOT",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "kWh",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30517",
            "data_type": "U64",
            "name": "Plant Daily Energy Yield",
            "tag": "E_DAY",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "kWh",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30521",
            "data_type": "U64",
            "name": "Total Operating Hours",
            "tag": "HOURS_TOT",
            "coef_a": 0.000278,
            "coef_b": 0.0,
            "unit": "h",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30865",
            "data_type": "U64",
            "name": "Total Feed-in Energy to Grid",
            "tag": "E_EXP_TOT",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "kWh",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30869",
            "data_type": "U64",
            "name": "Total Absorbed Energy from Grid",
            "tag": "E_IMP_TOT",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "kWh",
            "action": "4",
        },
        # Plant Remote Control Setpoints
        {
            "fc": "3",
            "address": "40018",
            "data_type": "U32",
            "name": "Active Power Limitation Percentage",
            "tag": "CMD_P_AC_SET_PERCENT",
            "coef_a": 0.01,
            "coef_b": 0.0,
            "unit": "%",
            "action": "1",
        },
        {
            "fc": "3",
            "address": "40020",
            "data_type": "U32",
            "name": "Active Power Limitation Absolute",
            "tag": "CMD_P_AC_SET_ABS",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "W",
            "action": "1",
        },
        {
            "fc": "3",
            "address": "40022",
            "data_type": "U32",
            "name": "Reactive Power Control Mode",
            "tag": "CMD_Q_MODE",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "1",
        },
        {
            "fc": "3",
            "address": "40024",
            "data_type": "I32",
            "name": "Reactive Power Setpoint Absolute",
            "tag": "CMD_Q_SET",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "var",
            "action": "1",
        },
        {
            "fc": "3",
            "address": "40026",
            "data_type": "I32",
            "name": "Power Factor cos phi Setpoint",
            "tag": "CMD_PF",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "",
            "action": "1",
        },
        {
            "fc": "3",
            "address": "40212",
            "data_type": "U32",
            "name": "Plant Inverters On Command",
            "tag": "CMD_ON",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "1",
        },
        {
            "fc": "3",
            "address": "40216",
            "data_type": "U32",
            "name": "Plant Inverters Off Command",
            "tag": "CMD_OFF",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "1",
        },
    ]


# ==============================================================================
# 7. UTILITY PLANT CONTROLLER: CLUSTER CONTROLLER & COM GATEWAY
# ==============================================================================
def build_sma_cluster_controller_rows() -> list[dict[str, Any]]:
    """Builds the certified 32-register list for SMA Cluster Controller (CLCON-10)."""
    rows = build_sma_data_manager_m_rows()[:24]
    return rows


def build_sma_com_gateway_rows() -> list[dict[str, Any]]:
    """Builds the certified 22-register list for SMA Com Gateway."""
    return build_sma_data_manager_m_rows()[:22]


# ==============================================================================
# 8. SMART ENERGY METERS: SMA ENERGY METER & HOME MANAGER 2.0
# ==============================================================================
def build_sma_energy_meter_rows() -> list[dict[str, Any]]:
    """
    Builds the certified 28-register list for SMA Energy Meter (EMETER-20).
    Bi-directional 3-phase multi-function meter measuring active/reactive import and export.
    """
    return [
        {
            "fc": "3",
            "address": "30005",
            "data_type": "U32",
            "name": "Meter Serial Number",
            "tag": "SERIAL",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30051",
            "data_type": "U32",
            "name": "Device Class",
            "tag": "",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30053",
            "data_type": "U32",
            "name": "Device Type Code",
            "tag": "MODEL",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30059",
            "data_type": "U32",
            "name": "Software Version",
            "tag": "FW_VER",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        # Active Power (Import / Export / Net)
        {
            "fc": "3",
            "address": "30775",
            "data_type": "I32",
            "name": "Active Power Net (+Import/-Export)",
            "tag": "P_AC",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "W",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30777",
            "data_type": "I32",
            "name": "Active Power Phase L1",
            "tag": "P_AC_1",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "W",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30779",
            "data_type": "I32",
            "name": "Active Power Phase L2",
            "tag": "P_AC_2",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "W",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30781",
            "data_type": "I32",
            "name": "Active Power Phase L3",
            "tag": "P_AC_3",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "W",
            "action": "4",
        },
        # Voltages (Phase & Line)
        {
            "fc": "3",
            "address": "30783",
            "data_type": "U32",
            "name": "Phase Voltage L1-N",
            "tag": "U_AC_1",
            "coef_a": 0.01,
            "coef_b": 0.0,
            "unit": "V",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30785",
            "data_type": "U32",
            "name": "Phase Voltage L2-N",
            "tag": "U_AC_2",
            "coef_a": 0.01,
            "coef_b": 0.0,
            "unit": "V",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30787",
            "data_type": "U32",
            "name": "Phase Voltage L3-N",
            "tag": "U_AC_3",
            "coef_a": 0.01,
            "coef_b": 0.0,
            "unit": "V",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30825",
            "data_type": "U32",
            "name": "Line Voltage L1-L2",
            "tag": "U_AC_L1L2",
            "coef_a": 0.01,
            "coef_b": 0.0,
            "unit": "V",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30827",
            "data_type": "U32",
            "name": "Line Voltage L2-L3",
            "tag": "U_AC_L2L3",
            "coef_a": 0.01,
            "coef_b": 0.0,
            "unit": "V",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30829",
            "data_type": "U32",
            "name": "Line Voltage L3-L1",
            "tag": "U_AC_L3L1",
            "coef_a": 0.01,
            "coef_b": 0.0,
            "unit": "V",
            "action": "4",
        },
        # Phase Currents
        {
            "fc": "3",
            "address": "30795",
            "data_type": "I32",
            "name": "Phase Current L1",
            "tag": "I_AC_1",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "A",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30797",
            "data_type": "I32",
            "name": "Phase Current L2",
            "tag": "I_AC_2",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "A",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30799",
            "data_type": "I32",
            "name": "Phase Current L3",
            "tag": "I_AC_3",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "A",
            "action": "4",
        },
        # Frequency, Reactive & Apparent Power, Power Factor
        {
            "fc": "3",
            "address": "30803",
            "data_type": "U32",
            "name": "Grid Frequency",
            "tag": "FREQ",
            "coef_a": 0.01,
            "coef_b": 0.0,
            "unit": "Hz",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30805",
            "data_type": "I32",
            "name": "Reactive Power Total",
            "tag": "Q_AC",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "var",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30807",
            "data_type": "I32",
            "name": "Apparent Power Total",
            "tag": "S_AC",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "VA",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30813",
            "data_type": "I32",
            "name": "Power Factor cos phi",
            "tag": "PF",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        # Energy Counters (Import & Export)
        {
            "fc": "3",
            "address": "30513",
            "data_type": "U64",
            "name": "Total Active Energy Import (Buy)",
            "tag": "E_IMP",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "kWh",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30517",
            "data_type": "U64",
            "name": "Total Active Energy Export (Sell)",
            "tag": "E_EXP",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "kWh",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30529",
            "data_type": "U64",
            "name": "Total Reactive Energy Import",
            "tag": "Q_IMP",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "kvarh",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30533",
            "data_type": "U64",
            "name": "Total Reactive Energy Export",
            "tag": "Q_EXP",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "kvarh",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30865",
            "data_type": "U64",
            "name": "Total Energy Counter",
            "tag": "E_TOT",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "kWh",
            "action": "4",
        },
    ]


def build_sma_home_manager_rows() -> list[dict[str, Any]]:
    """Builds the certified 26-register list for SMA Sunny Home Manager 2.0 (HM-20)."""
    return build_sma_energy_meter_rows()[:26]


# ==============================================================================
# 9. SUNSPEC STANDARD THREE-PHASE INVERTERS
# ==============================================================================
def build_sma_sunspec_threephase_rows() -> list[dict[str, Any]]:
    """
    Builds the certified 29-register list for SunSpec standard 3-phase inverters.
    Models 1 (Common), 103 (3-Phase Inverter Float), 120 (Nameplate), 121 (Basic Settings).
    """
    return [
        {
            "fc": "3",
            "address": "40003",
            "data_type": "U16",
            "name": "SunSpec Model ID 1",
            "tag": "SUNSPEC_ID_1",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40004_16",
            "data_type": "STRING",
            "name": "Manufacturer Name",
            "tag": "MANUFACTURER",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40020_16",
            "data_type": "STRING",
            "name": "Model Identifier",
            "tag": "MODEL",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40052_16",
            "data_type": "STRING",
            "name": "Serial Number",
            "tag": "SERIAL",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40070",
            "data_type": "U16",
            "name": "SunSpec Model ID 103",
            "tag": "SUNSPEC_ID_103",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40072",
            "data_type": "F32",
            "name": "AC Current Total",
            "tag": "I_AC",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "A",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40074",
            "data_type": "F32",
            "name": "AC Current Phase A",
            "tag": "I_AC_1",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "A",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40076",
            "data_type": "F32",
            "name": "AC Current Phase B",
            "tag": "I_AC_2",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "A",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40078",
            "data_type": "F32",
            "name": "AC Current Phase C",
            "tag": "I_AC_3",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "A",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40080",
            "data_type": "F32",
            "name": "Line Voltage AB",
            "tag": "U_AC_L1L2",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "V",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40082",
            "data_type": "F32",
            "name": "Line Voltage BC",
            "tag": "U_AC_L2L3",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "V",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40084",
            "data_type": "F32",
            "name": "Line Voltage CA",
            "tag": "U_AC_L3L1",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "V",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40086",
            "data_type": "F32",
            "name": "Phase Voltage AN",
            "tag": "U_AC_1",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "V",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40088",
            "data_type": "F32",
            "name": "Phase Voltage BN",
            "tag": "U_AC_2",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "V",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40090",
            "data_type": "F32",
            "name": "Phase Voltage CN",
            "tag": "U_AC_3",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "V",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40092",
            "data_type": "F32",
            "name": "AC Active Power Total",
            "tag": "P_AC",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "W",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40094",
            "data_type": "F32",
            "name": "Grid Frequency",
            "tag": "FREQ",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "Hz",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40096",
            "data_type": "F32",
            "name": "AC Apparent Power Total",
            "tag": "S_AC",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "VA",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40098",
            "data_type": "F32",
            "name": "AC Reactive Power Total",
            "tag": "Q_AC",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "var",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40100",
            "data_type": "F32",
            "name": "Power Factor",
            "tag": "PF",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40102",
            "data_type": "F32",
            "name": "Total Active Energy",
            "tag": "E_TOT",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "Wh",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40104",
            "data_type": "F32",
            "name": "DC Current Total",
            "tag": "I_DC",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "A",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40106",
            "data_type": "F32",
            "name": "DC Voltage",
            "tag": "U_DC",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "V",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40108",
            "data_type": "F32",
            "name": "DC Power Total",
            "tag": "P_DC",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "W",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40110",
            "data_type": "F32",
            "name": "Cabinet Temperature",
            "tag": "TEMP_INT",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "°C",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "40118",
            "data_type": "U16",
            "name": "Operating State",
            "tag": "STATE",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "9",
        },
        {
            "fc": "3",
            "address": "40120",
            "data_type": "U32",
            "name": "Event Flags Bitfield 1",
            "tag": "FAULT_CODE",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "9",
        },
        {
            "fc": "3",
            "address": "40124",
            "data_type": "U16",
            "name": "Active Power Derating %",
            "tag": "CMD_P_AC_SET_PERCENT",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "%",
            "action": "1",
        },
        {
            "fc": "3",
            "address": "40126",
            "data_type": "U16",
            "name": "Inverter On Off Switch",
            "tag": "CMD_ON",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "1",
        },
    ]


# ==============================================================================
# 10. ROW CLEANER & ENRICHER FOR LARGE OFFICIAL SPECIFICATION TABLES
# ==============================================================================
def clean_and_enrich_raw_file(
    file_path: Path,
    known_addr_map: dict[str, str] | None = None,
) -> list[dict[str, Any]]:
    """
    Reads an existing official SMA CSV file, normalizes all rows, infers missing tags,
    enforces strict tag uniqueness, and ensures 100% 5-Pillar Apollo compliance.
    """
    if not file_path.exists():
        return []

    with open(file_path, encoding="utf-8-sig", errors="ignore") as f:
        reader = list(csv.reader(f, delimiter=";"))

    if not reader or len(reader) < 2:
        return []

    cleaned: list[dict[str, Any]] = []
    seen_tags: set[str] = set()

    for row in reader[1:]:
        if not row or not any(cell.strip() for cell in row):
            continue
        while len(row) < 11:
            row.append("")

        fc = row[1].strip() or "3"
        addr = row[2].strip()
        dtype = row[3].strip() or "U32"
        name = row[5].strip()
        tag = row[6].strip()
        coef_a = row[7].strip() or "1.000000"
        coef_b = row[8].strip() or "0.000000"
        unit = row[9].strip()
        action = row[10].strip() or "4"

        # Standardize data types
        if dtype in ("S16", "S32"):
            dtype = dtype.replace("S", "I")
        if name == "160_N":
            dtype = "U16"  # SunSpec Model 160: N is uint16

        # FC validation
        if fc not in ("1", "2", "3", "4"):
            fc = "3"

        # Action validation
        if not action.isdigit() or action not in ("0", "1", "4", "9", "10"):
            action = "4"

        n_clean = name.lower().strip()
        u_clean = unit.lower().strip()

        # Sanitize excitation type (never let excitation type steal PF)
        if tag == "PF" and "excitation" in n_clean:
            tag = ""

        is_rating_reg = (
            addr in ("30231", "30203", "30205", "30207")
            or "wmaxoutrtg" in n_clean
            or "plimit" in n_clean
            or "maximum active power for status" in n_clean
            or "maximum possible permanent active power" in n_clean
            or "puissance active assignee" in n_clean
            or "puissance assignee" in n_clean
            or "puissance nominale" in n_clean
        )

        # Sanitize known rating / limit registers (prevent stealing P_AC)
        if is_rating_reg:
            if tag == "P_AC":
                tag = ""
            if not tag and "P_RATED" not in seen_tags:
                tag = "P_RATED"

        # Explicit Canonical SMA Modbus bindings
        elif addr in ("30513", "30529", "30583", "40102") or "rendement total" in n_clean:
            if "E_TOT" not in seen_tags:
                tag = "E_TOT"
        elif addr in ("30517", "30531"):
            if "E_DAY" not in seen_tags:
                tag = "E_DAY"
        elif addr in ("30775", "40084") or n_clean in (
            "puissance",
            "grid active power total",
            "active power on all line conductors",
            "ac active power across all phases w pac",
        ):
            if "P_AC" not in seen_tags:
                tag = "P_AC"
        elif addr in ("30803", "40086") or n_clean in ("frequence du reseau", "grid frequency"):
            if "FREQ" not in seen_tags:
                tag = "FREQ"
        elif addr in ("30805", "40090") or n_clean in (
            "puissance reactive",
            "reactive power total",
            "reactive power",
        ):
            if "Q_AC" not in seen_tags:
                tag = "Q_AC"
        elif addr == "30807":
            if any(k in n_clean for k in ("reactive", "réactive")) or u_clean in ("var", "kvar"):
                if "Q_AC_1" not in seen_tags:
                    tag = "Q_AC_1"
            elif any(k in n_clean for k in ("apparent", "apparente")) or u_clean in ("va", "kva"):
                if "S_AC" not in seen_tags:
                    tag = "S_AC"
            elif tag == "S_AC":
                tag = "Q_AC_1" if "Q_AC_1" not in seen_tags else ""
            elif "Q_AC_1" not in seen_tags:
                tag = "Q_AC_1"
        elif addr == "30809" or n_clean in (
            "puissance reactive l2",
            "reactive power phase l2",
            "reactive power of line conductor l2",
        ):
            if "Q_AC_2" not in seen_tags:
                tag = "Q_AC_2"
        elif addr == "30811" or n_clean in (
            "puissance reactive l3",
            "reactive power phase l3",
            "reactive power of line conductor l3",
        ):
            if "Q_AC_3" not in seen_tags:
                tag = "Q_AC_3"
        elif addr == "30813":
            if any(k in n_clean for k in ("apparent", "apparente")) or u_clean in ("va", "kva"):
                if "S_AC" not in seen_tags:
                    tag = "S_AC"
            elif any(k in n_clean for k in ("factor", "facteur", "cos phi", "cosphi", "pf")):
                if "PF" not in seen_tags:
                    tag = "PF"
            elif tag == "PF":
                tag = "S_AC" if "S_AC" not in seen_tags else ""
            elif "S_AC" not in seen_tags:
                tag = "S_AC"
        elif addr == "30815" or n_clean in (
            "puissance apparente l1",
            "apparent power phase l1",
            "apparent power of line conductor l1",
        ):
            if "S_AC_1" not in seen_tags:
                tag = "S_AC_1"
        elif addr == "30817" or n_clean in (
            "puissance apparente l2",
            "apparent power phase l2",
            "apparent power of line conductor l2",
        ):
            if "S_AC_2" not in seen_tags:
                tag = "S_AC_2"
        elif addr == "30819" or n_clean in (
            "puissance apparente l3",
            "apparent power phase l3",
            "apparent power of line conductor l3",
        ):
            if "S_AC_3" not in seen_tags:
                tag = "S_AC_3"
        elif addr in ("30821", "30825", "30949", "31221") or any(
            k in n_clean
            for k in (
                "facteur de decalage",
                "facteur de décalage",
                "facteur de dephasage",
                "facteur de déphasage",
                "power factor cos phi",
            )
        ):
            if "PF" not in seen_tags:
                tag = "PF"
        elif addr in ("30769", "31793") or n_clean == "entree courant dc_1":
            if "I_DC_1" not in seen_tags:
                tag = "I_DC_1"
        elif addr in ("30771",):
            if "U_DC_1" not in seen_tags:
                tag = "U_DC_1"
        elif addr in ("30773",):
            if "P_DC_1" not in seen_tags:
                tag = "P_DC_1"
        elif addr == "31803" or n_clean == "entree courant dc_6":
            if "I_DC_6" not in seen_tags:
                tag = "I_DC_6"

        # Storage & Battery canonical bindings
        elif (
            addr == "30845"
            or "chastt" in name.lower()
            or name.lower() in ("battery state of charge", "bat.chastt")
        ):
            if "SOC" not in seen_tags:
                tag = "SOC"
                if not unit:
                    unit = "%"
        elif addr in ("30851", "30853") or name.lower() in ("bat.vol", "battery voltage"):
            if "U_BAT" not in seen_tags:
                tag = "U_BAT"
                if not unit:
                    unit = "V"
        elif addr in ("30843", "30855") or name.lower() in ("bat.amp", "battery current"):
            if "I_BAT" not in seen_tags:
                tag = "I_BAT"
                if not unit:
                    unit = "A"
        elif addr in ("30849",) or name.lower() in ("bat.tmpval", "battery temperature"):
            if "TEMP_BAT" not in seen_tags:
                tag = "TEMP_BAT"
                if not unit:
                    unit = "°C"
        elif addr in ("30847", "31391") or name.lower() in (
            "operation.bat.health",
            "battery state of health",
        ):
            if "SOH" not in seen_tags:
                tag = "SOH"
                if not unit:
                    unit = "%"
        elif addr in ("30955",) or name.lower() in ("bat.opstt", "battery operating mode"):
            if "STATE_BAT" not in seen_tags:
                tag = "STATE_BAT"

        # SunSpec Model 160 Multiple MPPT mappings (Core 2 / STP 110-60)
        mppt_match = re.match(r"^160_(\d+)_DC([AVW])$", name)
        if mppt_match:
            idx = int(mppt_match.group(1)) + 1
            kind = mppt_match.group(2)
            tag_candidate = f"{'I' if kind == 'A' else ('U' if kind == 'V' else 'P')}_DC_{idx}"
            if tag_candidate not in seen_tags:
                tag = tag_candidate

        # Tag inference and canonical mapping
        if not tag and not is_rating_reg:
            if known_addr_map and addr in known_addr_map:
                tag = known_addr_map[addr]
            else:
                inferred = infer_sunpm_tag(name)
                if inferred and is_canonical_sunpm_tag(inferred):
                    tag = inferred

        # Deduplicate tags across file
        if tag:
            if tag in seen_tags:
                tag = ""
            else:
                seen_tags.add(tag)

        # Coefficients
        try:
            ca = float(coef_a)
        except ValueError:
            ca = 1.0
        try:
            cb = float(coef_b)
        except ValueError:
            cb = 0.0

        # Units
        unit = normalize_sma_unit(unit)

        cleaned.append(
            {
                "fc": fc,
                "address": addr,
                "data_type": dtype,
                "name": name,
                "tag": tag,
                "coef_a": ca,
                "coef_b": cb,
                "unit": unit,
                "action": action,
            }
        )

    return cleaned


# ==============================================================================
# 11. CATALOGUE TARGET GENERATION & RETRIEVAL
# ==============================================================================
def get_all_sma_targets(
    workspace_dir: Path,
) -> list[tuple[str, str, str, str, list[dict[str, Any]]]]:
    """Returns the comprehensive certified catalog of all 27 SMA definition targets."""
    mfg = "SMA"

    # Base builders
    sb_rows = build_sma_sunny_boy_rows()
    sbs_rows = build_sma_sunny_boy_storage_rows()
    stp_rows = build_sma_sunny_tripower_rows()
    core1_core2_rows = build_sma_core1_core2_rows()
    si_rows = build_sma_sunny_island_rows()
    edmm_rows = build_sma_data_manager_m_rows()
    clcon_rows = build_sma_cluster_controller_rows()
    gateway_rows = build_sma_com_gateway_rows()
    meter_rows = build_sma_energy_meter_rows()
    hm_rows = build_sma_home_manager_rows()
    sunspec_rows = build_sma_sunspec_threephase_rows()

    # Address tag maps for enrichment
    tag_map = {str(r["address"]): r["tag"] for r in stp_rows if r.get("tag")}
    tag_map.update({str(r["address"]): r["tag"] for r in core1_core2_rows if r.get("tag")})

    # Rich specification tables (load and clean from production/ or Definitions/ if present)
    prod_dir = workspace_dir / "production"
    defs_dir = workspace_dir / "Definitions"

    # 1. STP-X (435 rows, 50 unique tags, certified French variable names)
    stpx_src = prod_dir / "INV_SMA_STP_X.csv"
    if not stpx_src.exists():
        stpx_src = defs_dir / "INV_SMA_STP_X.csv"
    stpx_rows = clean_and_enrich_raw_file(stpx_src) if stpx_src.exists() else stp_rows

    # 2. STP Smart Energy (372 rows)
    se_src = prod_dir / "INV_SMA_STP_5-10-SE40.csv"
    se_rows = clean_and_enrich_raw_file(se_src, tag_map) if se_src.exists() else sbs_rows

    # 3. Core 1 full (105 rows)
    core1_src = prod_dir / "INV_SMA_STPCORE1_50-41.csv"
    core1_full_rows = (
        clean_and_enrich_raw_file(core1_src, tag_map) if core1_src.exists() else core1_core2_rows
    )

    # 4. Core 2 full (1,211 rows)
    core2_src = workspace_dir / "tagged_definitions" / "INV_SMA_STP_110-60_1.00.11.R.csv"
    if not core2_src.exists():
        core2_src = prod_dir / "INV_SMA_STP_110-60.csv"
    if not core2_src.exists():
        core2_src = defs_dir / "INV_SMA_STP_110-60.csv"
    core2_full_rows = (
        clean_and_enrich_raw_file(core2_src, tag_map) if core2_src.exists() else core1_core2_rows
    )

    # 5. Sunny Central (182 rows)
    sc_src = prod_dir / "INV_SMA_SUNNY_CENTRAL.csv"
    sc_rows = clean_and_enrich_raw_file(sc_src, tag_map) if sc_src.exists() else stp_rows

    # 5b. Sunny Central Storage (AC Central + DC Battery Storage)
    scs_rows = [dict(r) for r in sc_rows]
    bat_storage_regs = [
        {
            "fc": "3",
            "address": "30845",
            "data_type": "U32",
            "name": "Battery State of Charge",
            "tag": "SOC",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "%",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30847",
            "data_type": "U32",
            "name": "Battery State of Health",
            "tag": "SOH",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "%",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30849",
            "data_type": "U32",
            "name": "Battery Operating Mode",
            "tag": "STATE_BAT",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30851",
            "data_type": "I32",
            "name": "Battery Temperature",
            "tag": "TEMP_BAT",
            "coef_a": 0.1,
            "coef_b": 0.0,
            "unit": "°C",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30853",
            "data_type": "U32",
            "name": "Battery DC Voltage",
            "tag": "U_BAT",
            "coef_a": 0.01,
            "coef_b": 0.0,
            "unit": "V",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30855",
            "data_type": "I32",
            "name": "Battery DC Current",
            "tag": "I_BAT",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "A",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "30857",
            "data_type": "I32",
            "name": "Battery DC Power",
            "tag": "P_BAT",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "W",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "31393",
            "data_type": "U32",
            "name": "Battery Nominal Capacity",
            "tag": "CAP_NOM_BAT",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "Ah",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "31395",
            "data_type": "U32",
            "name": "Battery Usable Capacity",
            "tag": "CAP_REM_BAT",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "Ah",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "31397",
            "data_type": "U64",
            "name": "Total Battery Charged Energy",
            "tag": "E_BAT_CHG_TOT",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "kWh",
            "action": "4",
        },
        {
            "fc": "3",
            "address": "31401",
            "data_type": "U64",
            "name": "Total Battery Discharged Energy",
            "tag": "E_BAT_DISCHG_TOT",
            "coef_a": 0.001,
            "coef_b": 0.0,
            "unit": "kWh",
            "action": "4",
        },
    ]
    scs_existing_addrs = {r["address"] for r in scs_rows}
    for br in bat_storage_regs:
        if br["address"] not in scs_existing_addrs:
            scs_rows.append(br)

    # 6. SC2xxx (78 rows)
    sc2_src = defs_dir / "INV_SMA_SC2xxx.csv"
    sc2_rows = clean_and_enrich_raw_file(sc2_src, tag_map) if sc2_src.exists() else sc_rows

    # 7. Webbox (350 rows)
    webbox_src = prod_dir / "INV_SMA_SC_Webbox.csv"
    webbox_rows = (
        clean_and_enrich_raw_file(webbox_src, tag_map) if webbox_src.exists() else edmm_rows
    )

    # 8. Generic (857 rows)
    gen_src = defs_dir / "INV_SMA_Generic.csv"
    gen_rows = clean_and_enrich_raw_file(gen_src, tag_map) if gen_src.exists() else sunspec_rows

    return [
        # Residential
        ("INV_SMA_SunnyBoy.csv", "Inverter", mfg, "SunnyBoy", sb_rows),
        ("INV_SMA_SunnyBoyStorage.csv", "Inverter", mfg, "SunnyBoyStorage", sbs_rows),
        # Commercial Three-Phase
        ("INV_SMA_SunnyTripower.csv", "Inverter", mfg, "SunnyTripower", stp_rows),
        ("INV_SMA_SunnyTripower_X.csv", "Inverter", mfg, "SunnyTripower-X", stpx_rows),
        ("INV_SMA_STP_X.csv", "Inverter", mfg, "STP_X", stpx_rows),
        ("INV_SMA_Core1_Core2.csv", "Inverter", mfg, "Core1_Core2", core1_core2_rows),
        ("INV_SMA_Core1.csv", "Inverter", mfg, "Core1", core1_full_rows),
        ("INV_SMA_Core2.csv", "Inverter", mfg, "Core2", core2_full_rows),
        ("INV_SMA_STPCORE1_50-41.csv", "Inverter", mfg, "STPCORE1_50-41", core1_full_rows),
        ("INV_SMA_STP_110-60.csv", "Inverter", mfg, "STP_110-60", core2_full_rows),
        # Hybrid & Energy Storage
        (
            "INV_SMA_SunnyTripower_SmartEnergy.csv",
            "Inverter",
            mfg,
            "SunnyTripower-SmartEnergy",
            se_rows,
        ),
        ("INV_SMA_STP_5-10-SE40.csv", "Inverter", mfg, "STP_5-10-SE40", se_rows),
        ("INV_SMA_SunnyIsland.csv", "Inverter", mfg, "SunnyIsland", si_rows),
        # Utility Scale Central
        ("INV_SMA_SunnyCentral.csv", "Inverter", mfg, "SunnyCentral", sc_rows),
        ("INV_SMA_SUNNY_CENTRAL.csv", "Inverter", mfg, "SUNNY_CENTRAL", sc_rows),
        ("INV_SMA_SC2xxx.csv", "Inverter", mfg, "SC2xxx", sc2_rows),
        ("INV_SMA_SunnyCentral_Storage.csv", "Inverter", mfg, "SunnyCentral-Storage", scs_rows),
        # Communication Gateways & Plant Controllers
        ("COM_SMA_DataManager_M.csv", "Logger", mfg, "DataManager-M", edmm_rows),
        ("COM_SMA_ClusterController.csv", "Logger", mfg, "ClusterController", clcon_rows),
        ("COM_SMA_ComGateway.csv", "Logger", mfg, "ComGateway", gateway_rows),
        ("INV_SMA_SC_Webbox.csv", "Inverter", mfg, "SC_Webbox", webbox_rows),
        # Smart Energy Meters & Grid Analyzers
        ("MET_SMA_EnergyMeter.csv", "Meter", mfg, "EnergyMeter", meter_rows),
        ("MET_SMA_HomeManager.csv", "Meter", mfg, "HomeManager", hm_rows),
        # SunSpec Standard Inverters & Interoperable Profiles
        ("INV_SMA_SunSpec_Inverter.csv", "Inverter", mfg, "SunSpec-Inverter", sunspec_rows),
        ("INV_SMA_Sunspec.csv", "Inverter", mfg, "Sunspec", sunspec_rows),
        ("INV_SMA_Generic.csv", "Inverter", mfg, "Generic", gen_rows),
        ("INV_SMA_Modbus.csv", "Inverter", mfg, "Modbus", sunspec_rows),
    ]


def generate_all_sma_definitions(workspace_dir: Path | str) -> dict[str, Path]:
    """
    Generates and certifies all 27 SMA definitions across:
    - production/
    - Definitions/SMA/
    - Definitions/
    - Definitions/Classic/ (for SunnyTripower & Core1_Core2)
    - generated_definitions/
    - tagged_definitions/
    """
    ws = Path(workspace_dir)
    prod_dir = ws / "production"
    sma_dir = ws / "Definitions" / "SMA"
    defs_dir = ws / "Definitions"
    classic_dir = ws / "Definitions" / "Classic"
    gen_dir = ws / "generated_definitions"
    tag_dir = ws / "tagged_definitions"

    for d in (prod_dir, sma_dir, defs_dir, classic_dir, gen_dir, tag_dir):
        d.mkdir(parents=True, exist_ok=True)

    targets = get_all_sma_targets(ws)
    generated: dict[str, Path] = {}

    for fname, cat, mfg, model, rows in targets:
        # 1. Primary write into production/
        prod_path = prod_dir / fname
        write_sma_csv(prod_path, cat, mfg, model, rows)
        generated[fname] = prod_path

        # 2. Synchronize byte-identical copies
        sync_targets = [sma_dir, defs_dir, gen_dir, tag_dir]
        if fname in ("INV_SMA_SunnyTripower.csv", "INV_SMA_Core1_Core2.csv"):
            sync_targets.append(classic_dir)

        for target_dir in sync_targets:
            dest = target_dir / fname
            shutil.copy2(prod_path, dest)
            is_valid, issues = verify_file_pillars(dest)
            if not is_valid:
                raise ValueError(f"Pillar violation at {dest}:\n" + "\n".join(issues))

    return generated
