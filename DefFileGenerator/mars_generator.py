#!/usr/bin/env python3
"""
Dedicated Mars Renewable Definition Generator
Strictly adheres to WebdynSunPM standard specification and user instructions:
- MarsRenewable_LocalControler.csv (ID 247: LC Overview FC4 + EMS Writes FC3)
- MarsRenewable_PCS-Bank-Rack.csv (Device ID 1: FC2 Alarms + DIO1 + PCS1 + Bank1)
- MarsRenewable_Rack-Cell.csv (Device ID 2: Rack 1 + Cell 1..N)
- MarsRenewable_LC-AC.csv (Device ID 1: Liquid Cooling 1 + Air Conditioning 1)
- MarsRenewable_Meter.csv (Device ID 1: Bidirectional Meter 1)
- MarsRenewable_PV.csv (Device ID 1: PV Inverter 1)
"""

import csv
from pathlib import Path
from typing import Any

import openpyxl

from DefFileGenerator.def_gen import Generator


def normalize_unit(unit_raw: Any) -> tuple[str, float]:
    """Normalizes engineering units and extracts scaling factors."""
    if not unit_raw:
        return "", 1.0
    u = str(unit_raw).strip()
    if u in ("-", "/", "1"):
        return "", 1.0
    if u.lower() == "0.01hz":
        return "Hz", 0.01
    if u.lower() in ("℃", "°c", "degc", "celsius", "degr.celsius"):
        return "°C", 1.0
    if u.lower() in ("ω", "ohm", "ohms"):
        return "Ω", 1.0
    if u.lower() == "kvar":
        return "kvar", 1.0
    if u.lower() == "var":
        return "Var", 1.0
    return u, 1.0


def normalize_type(type_raw: Any) -> str:
    """Maps vendor type strings to canonical WebdynSunPM data types."""
    if not type_raw:
        return "U16"
    t = str(type_raw).strip().lower()
    if "float" in t or t in ("f32", "float32"):
        return "F32"
    if t in ("u32", "uint32"):
        return "U32"
    if t in ("i32", "int32"):
        return "I32"
    if t in ("i16", "int16"):
        return "I16"
    if t in ("u16", "uint16", "bit", "bits", "bool", "boolean"):
        return "U16"
    return "U16"


def determine_action(name: str, fc: str, is_reserve: bool, default_action: str = "4") -> str:
    """Determines WebdynSunPM Action code based on signal role."""
    if is_reserve:
        return "0"
    if fc in ("1", "16", "3") and (
        "mode" in name.lower()
        or "duration" in name.lower()
        or ("power" in name.lower() and "set" in name.lower())
        or "working mode" in name.lower()
    ):
        return "1"
    name_lower = name.lower()
    if any(k in name_lower for k in ("alarm", "status", "fault", "crash-stop", "trip", "feedback", "indicator")):
        return "9"
    return default_action


def write_definition_file(
    output_path: str | Path,
    protocol: str,
    category: str,
    manufacturer: str,
    model: str,
    registers: list[dict[str, Any]],
    strict_overlap: bool = True,
) -> tuple[bool, Any]:
    """Writes a strictly compliant WebdynSunPM definition file with UTF-8 BOM encoding."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, delimiter=";", lineterminator="\r\n")
        # 11 columns in header (10 semicolons)
        writer.writerow([protocol, category, manufacturer, model, "", "", "", "", "", "", ""])

        for idx, reg in enumerate(registers, start=1):
            writer.writerow(
                [
                    str(idx),
                    str(reg["Info1"]),
                    str(reg["Info2"]),
                    str(reg["Info3"]),
                    "",
                    str(reg["Name"]),
                    str(reg.get("Tag", "")),
                    f"{float(reg['CoefA']):.6f}",
                    f"{float(reg['CoefB']):.6f}",
                    str(reg.get("Unit", "")),
                    str(reg["Action"]),
                ]
            )

    gen = Generator()
    report = gen.validate_csv_detailed(str(output_path), strict=True, strict_overlap=strict_overlap)
    return report.is_valid, report


def parse_mars_workbook_all(xlsx_path: str | Path) -> dict[str, list[dict[str, Any]]]:
    """Parses all required sheets from the Mars Local Controller V7.0 workbook."""
    wb = openpyxl.load_workbook(str(xlsx_path), data_only=True)

    # 1. Sheet: LC Information
    ws_lc = wb["LC Information"]
    lc_rows = list(ws_lc.iter_rows(values_only=True))

    # Overview: FC 0x04, ID 247 (rows 6 to 10)
    lc_overview = []
    for r in lc_rows[5:10]:
        if r and r[0] is not None and str(r[0]).strip().isdigit():
            addr = int(str(r[0]).strip())
            name = f"LC Overview - {str(r[2]).strip()}"
            dtype = normalize_type(r[4])
            lc_overview.append(
                {
                    "Info1": "4",
                    "Info2": str(addr),
                    "Info3": dtype,
                    "Name": name,
                    "Tag": "",
                    "CoefA": 1.0,
                    "CoefB": 0.0,
                    "Unit": "",
                    "Action": "4",
                }
            )

    # EMS Writes: FC 0x16 / FC 0x03, ID 247 (rows 18 to 24)
    lc_writes = []
    ems_tags = {
        0: "",
        1: "",
        3: "",
        4: "CMD_P_AC_SET_ABS",
        6: "",
        8: "",
        9: "",
    }
    for r in lc_rows[17:24]:
        if r and r[0] is not None and str(r[0]).strip().isdigit():
            addr = int(str(r[0]).strip())
            name = f"EMS - {str(r[2]).strip()}"
            unit, factor = normalize_unit(r[3])
            dtype = normalize_type(r[4])
            # Reactive power is marked as not supported / empty point in notes
            action = "0" if addr == 6 else "1"
            tag = ems_tags.get(addr, "")
            lc_writes.append(
                {
                    "Info1": "3",
                    "Info2": str(addr),
                    "Info3": dtype,
                    "Name": name,
                    "Tag": tag,
                    "CoefA": factor,
                    "CoefB": 0.0,
                    "Unit": unit,
                    "Action": action,
                }
            )

    # LC Alarms: FC 0x02, ID 1 (rows 30 to 33)
    lc_alarms = []
    for r in lc_rows[29:33]:
        if r and r[0] is not None and str(r[0]).strip().isdigit():
            addr = int(str(r[0]).strip())
            name = f"{str(r[2]).strip()} [0=normal, 1=alarm]"
            lc_alarms.append(
                {
                    "Info1": "2",
                    "Info2": str(addr),
                    "Info3": "U16",
                    "Name": name,
                    "Tag": "",
                    "CoefA": 1.0,
                    "CoefB": 0.0,
                    "Unit": "",
                    "Action": "9",
                }
            )

    # Helper to parse single equipment template
    def parse_single_equipment(sheet_name: str, skip_header: int = 6) -> list[dict[str, Any]]:
        ws = wb[sheet_name]
        items = []
        rows = list(ws.iter_rows(values_only=True))
        for r in rows[skip_header:]:
            if not r or r[0] is None:
                continue

            # Check if any cell has '...' indicating equipment 1 section has ended
            row_str = " ".join(str(c) for c in r if c is not None)
            if "..." in row_str:
                break

            equip = str(r[1]).strip() if len(r) > 1 and r[1] else ""
            if equip and not equip.startswith("1#"):
                break

            addr_str = str(r[0]).strip()
            if not addr_str.isdigit():
                continue
            addr = int(addr_str)

            name = str(r[2]).strip() if len(r) > 2 and r[2] else ""
            acc = str(r[3]).strip() if len(r) > 3 and r[3] else ""
            unit_val = str(r[4]).strip() if len(r) > 4 and r[4] else ""
            type_val = str(r[5]).strip() if len(r) > 5 and r[5] else ""
            notes = str(r[6]).strip() if len(r) > 6 and r[6] else ""

            unit, unit_factor = normalize_unit(unit_val)
            acc_factor = 1.0
            if acc:
                try:
                    acc_factor = float(acc)
                except ValueError:
                    pass
            coef_a = acc_factor if acc_factor != 1.0 else unit_factor
            dtype = normalize_type(type_val)
            is_reserve = "reserve" in name.lower()
            action = determine_action(name, "4", is_reserve)

            # Fix known vendor overlap typo: Branch 2~n DC Current listed at 71 instead of 72
            if addr == 71 and "branch 2~n dc current" in name.lower():
                addr = 72

            items.append(
                {
                    "addr": addr,
                    "name": name,
                    "dtype": dtype,
                    "unit": unit,
                    "coef_a": coef_a,
                    "action": action,
                    "notes": notes,
                }
            )
        return items

    pcs_items = parse_single_equipment("PCS Information")
    bank_items = parse_single_equipment("Bank Information")
    meter_items = parse_single_equipment("Meter Information")
    ac_items = parse_single_equipment("Air Conditioning Information")
    dio_items = parse_single_equipment("DIO Information")
    pv_items = parse_single_equipment("PV Information")
    lc_cool_items = parse_single_equipment("Liquid Cooling Information")

    # Rack and Cell Information
    ws_rack = wb["Rack and Cell Information"]
    rack_rows = list(ws_rack.iter_rows(values_only=True))
    rack_items = []
    cell_items = []
    in_cell = False

    for r in rack_rows[6:]:
        if not r or r[0] is None:
            continue
        addr_str = str(r[0]).strip()
        if not addr_str.isdigit():
            if "..." in addr_str:
                in_cell = True
            continue
        addr = int(addr_str)
        name = str(r[2]).strip() if len(r) > 2 and r[2] else ""
        unit_val = str(r[4]).strip() if len(r) > 4 and r[4] else ""
        type_val = str(r[5]).strip() if len(r) > 5 and r[5] else ""
        unit, factor = normalize_unit(unit_val)
        dtype = normalize_type(type_val)
        action = "9" if "status" in name.lower() or "alarm" in name.lower() else "4"

        if addr < 100 and not in_cell:
            rack_items.append(
                {
                    "addr": addr,
                    "name": name,
                    "dtype": dtype,
                    "unit": unit,
                    "coef_a": factor,
                    "action": action,
                }
            )
        elif 100 <= addr <= 109:
            cell_items.append(
                {
                    "addr": addr,
                    "name": name,
                    "dtype": dtype,
                    "unit": unit,
                    "coef_a": factor,
                    "action": action,
                }
            )

    return {
        "lc_overview": lc_overview,
        "lc_writes": lc_writes,
        "lc_alarms": lc_alarms,
        "pcs": pcs_items,
        "bank": bank_items,
        "meter": meter_items,
        "ac": ac_items,
        "dio": dio_items,
        "pv": pv_items,
        "lc_cooling": lc_cool_items,
        "rack": rack_items,
        "cell": cell_items,
    }


def generate_mars_requested_definitions(
    xlsx_path: str | Path,
    output_dir: str | Path,
    num_cells: int = 1,
) -> dict[str, Path]:
    """
    Generates the exact 6 Mars Renewable definition files as requested.
    Returns dictionary of generated file paths.
    """
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    raw = parse_mars_workbook_all(xlsx_path)
    generated = {}

    # -------------------------------------------------------------------------
    # 1. MarsRenewable_LocalControler.csv (ID 247)
    # -------------------------------------------------------------------------
    lc_regs = []
    # Overview (FC4)
    for reg in raw["lc_overview"]:
        lc_regs.append(dict(reg))
    # EMS writes (FC3)
    for reg in raw["lc_writes"]:
        lc_regs.append(dict(reg))

    p1 = out_dir / "MarsRenewable_LocalControler.csv"
    valid, rep = write_definition_file(
        p1, "modbusTCP", "Inverter", "MarsRenewable", "LocalControler", lc_regs, strict_overlap=True
    )
    if not valid:
        raise ValueError(f"Validation failed for {p1.name}: {rep.issues}")
    generated["MarsRenewable_LocalControler.csv"] = p1

    # -------------------------------------------------------------------------
    # 2. MarsRenewable_PCS-Bank-Rack.csv (Device ID 1: Alarms + DIO1 + PCS1 + Bank1)
    # -------------------------------------------------------------------------
    pcs_bank_regs = []
    # FC2 Alarms
    for reg in raw["lc_alarms"]:
        pcs_bank_regs.append(dict(reg))

    # DIO1 (Addr 20006..20021, strictly 1#DIO channels)
    for item in raw["dio"]:
        if item["action"] != "0" and item["addr"] <= 20021:
            clean_name = item["name"].replace("（", " (").replace("）", ")")
            pcs_bank_regs.append(
                {
                    "Info1": "4",
                    "Info2": str(item["addr"]),
                    "Info3": item["dtype"],
                    "Name": f"DIO1 {clean_name}",
                    "Tag": "",
                    "CoefA": item["coef_a"],
                    "CoefB": 0.0,
                    "Unit": item["unit"],
                    "Action": item["action"],
                }
            )

    # PCS1 (Addr 0..60, plus branches 1..6: Addr 62..96)
    pcs_tag_map = {
        0: "",
        2: "U_AC_L1L2",
        4: "U_AC_L2L3",
        6: "U_AC_L3L1",
        8: "I_AC_1",
        10: "I_AC_2",
        12: "I_AC_3",
        14: "FREQ",
        16: "P_AC",
        18: "Q_AC",
        20: "S_AC",
        22: "PF",
        24: "P_AC_1",
        26: "P_AC_2",
        28: "P_AC_3",
        30: "Q_AC_1",
        32: "Q_AC_2",
        34: "Q_AC_3",
        36: "S_AC_1",
        38: "S_AC_2",
        40: "S_AC_3",
        42: "PF_1",
        44: "PF_2",
        46: "PF_3",
        48: "TEMP_INT",
        50: "TEMP_AMB",
        54: "E_TOT",
        58: "E_DAY",
    }
    used_tags: set[str] = set()

    for item in raw["pcs"]:
        if item["action"] != "0":
            addr = item["addr"]
            # Skip the single template row 'Branch 2~n' as we expand branches 1..6 below
            if addr in (68, 70, 72) and "2~n" in item["name"]:
                continue

            tag = pcs_tag_map.get(addr, "")
            if tag and tag in used_tags:
                tag = ""
            if tag:
                used_tags.add(tag)

            pcs_bank_regs.append(
                {
                    "Info1": "4",
                    "Info2": str(addr),
                    "Info3": item["dtype"],
                    "Name": f"PCS1 {item['name']}",
                    "Tag": tag,
                    "CoefA": item["coef_a"],
                    "CoefB": 0.0,
                    "Unit": item["unit"],
                    "Action": item["action"],
                }
            )

    # Expand PCS Branches 2 to 6 (stride 6, 68..96)
    for b in range(2, 7):
        b_base = 62 + (b - 1) * 6
        pcs_bank_regs.append(
            {
                "Info1": "4",
                "Info2": str(b_base),
                "Info3": "F32",
                "Name": f"PCS1 Branch {b} DC Power",
                "Tag": "",
                "CoefA": 1.0,
                "CoefB": 0.0,
                "Unit": "kW",
                "Action": "4",
            }
        )
        pcs_bank_regs.append(
            {
                "Info1": "4",
                "Info2": str(b_base + 2),
                "Info3": "F32",
                "Name": f"PCS1 Branch {b} DC Voltage",
                "Tag": "",
                "CoefA": 1.0,
                "CoefB": 0.0,
                "Unit": "V",
                "Action": "4",
            }
        )
        pcs_bank_regs.append(
            {
                "Info1": "4",
                "Info2": str(b_base + 4),
                "Info3": "F32",
                "Name": f"PCS1 Branch {b} DC Current",
                "Tag": "",
                "CoefA": 1.0,
                "CoefB": 0.0,
                "Unit": "A",
                "Action": "4",
            }
        )

    # Bank1 (Addr 5000..5024)
    bank_tag_map = {
        5000: "U_DC",
        5002: "I_DC",
        5004: "SOC",
        5006: "SOH",
        5008: "R_ISO",
    }
    for item in raw["bank"]:
        if item["action"] != "0":
            addr = item["addr"]
            tag = bank_tag_map.get(addr, "")
            if tag and tag in used_tags:
                tag = ""
            if tag:
                used_tags.add(tag)

            pcs_bank_regs.append(
                {
                    "Info1": "4",
                    "Info2": str(addr),
                    "Info3": item["dtype"],
                    "Name": f"Bank1 {item['name']}",
                    "Tag": tag,
                    "CoefA": item["coef_a"],
                    "CoefB": 0.0,
                    "Unit": item["unit"],
                    "Action": item["action"],
                }
            )

    p2 = out_dir / "MarsRenewable_PCS-Bank-Rack.csv"
    valid, rep = write_definition_file(
        p2, "modbusTCP", "Battery", "MarsRenewable", "PCS-Bank", pcs_bank_regs, strict_overlap=True
    )
    if not valid:
        raise ValueError(f"Validation failed for {p2.name}: {rep.issues}")
    generated["MarsRenewable_PCS-Bank-Rack.csv"] = p2

    # -------------------------------------------------------------------------
    # 3. MarsRenewable_Rack-Cell.csv (Rack 1 + Cells)
    # -------------------------------------------------------------------------
    rack_cell_regs = []
    rack_tag_map = {
        1: "U_DC",
        3: "I_DC",
        5: "SOC",
        7: "SOH",
        9: "R_ISO",
        25: "TEMP_CELL",
    }
    used_rack_tags: set[str] = set()

    for item in raw["rack"]:
        if item["action"] != "0":
            addr = item["addr"]
            tag = rack_tag_map.get(addr, "")
            if tag and tag in used_rack_tags:
                tag = ""
            if tag:
                used_rack_tags.add(tag)

            rack_cell_regs.append(
                {
                    "Info1": "4",
                    "Info2": str(addr),
                    "Info3": item["dtype"],
                    "Name": f"Rack {item['name']}",
                    "Tag": tag,
                    "CoefA": item["coef_a"],
                    "CoefB": 0.0,
                    "Unit": item["unit"],
                    "Action": item["action"],
                }
            )

    # Add Cell 1 (and optional up to num_cells)
    cell_template = raw["cell"]
    for c_idx in range(1, num_cells + 1):
        cell_base = 100 + (c_idx - 1) * 10
        for item in cell_template:
            c_rel = item["addr"] - 100
            c_addr = cell_base + c_rel
            rack_cell_regs.append(
                {
                    "Info1": "4",
                    "Info2": str(c_addr),
                    "Info3": item["dtype"],
                    "Name": f"Cell {c_idx} {item['name']}",
                    "Tag": "",  # Cell level measurements have empty canonical tags to prevent collisions
                    "CoefA": item["coef_a"],
                    "CoefB": 0.0,
                    "Unit": item["unit"],
                    "Action": item["action"],
                }
            )

    p3 = out_dir / "MarsRenewable_Rack-Cell.csv"
    valid, rep = write_definition_file(
        p3, "modbusTCP", "Battery", "MarsRenewable", "Rack-Cell", rack_cell_regs, strict_overlap=True
    )
    if not valid:
        raise ValueError(f"Validation failed for {p3.name}: {rep.issues}")
    generated["MarsRenewable_Rack-Cell.csv"] = p3

    # -------------------------------------------------------------------------
    # 4. MarsRenewable_LC-AC.csv (Liquid Cooling 1 + Air Conditioning 1)
    # -------------------------------------------------------------------------
    lc_ac_regs = []
    used_cooling_tags: set[str] = set()

    # Liquid Cooling (Addr 30000..30032)
    lc_tag_map = {
        30006: "TEMP_AMB",
    }
    for item in raw["lc_cooling"]:
        if item["action"] != "0":
            addr = item["addr"]
            tag = lc_tag_map.get(addr, "")
            if tag and tag in used_cooling_tags:
                tag = ""
            if tag:
                used_cooling_tags.add(tag)

            lc_ac_regs.append(
                {
                    "Info1": "4",
                    "Info2": str(addr),
                    "Info3": item["dtype"],
                    "Name": f"LiquidCooling1 {item['name']}",
                    "Tag": tag,
                    "CoefA": item["coef_a"],
                    "CoefB": 0.0,
                    "Unit": item["unit"],
                    "Action": item["action"],
                }
            )

    # Air Conditioning (Addr 15000..15027)
    ac_tag_map = {
        15020: "TEMP",
        15027: "HUMIDITY",
    }
    for item in raw["ac"]:
        if item["action"] != "0":
            addr = item["addr"]
            tag = ac_tag_map.get(addr, "")
            if tag and tag in used_cooling_tags:
                tag = ""
            if tag:
                used_cooling_tags.add(tag)

            lc_ac_regs.append(
                {
                    "Info1": "4",
                    "Info2": str(addr),
                    "Info3": item["dtype"],
                    "Name": f"AC1 {item['name']}",
                    "Tag": tag,
                    "CoefA": item["coef_a"],
                    "CoefB": 0.0,
                    "Unit": item["unit"],
                    "Action": item["action"],
                }
            )

    p4 = out_dir / "MarsRenewable_LC-AC.csv"
    valid, rep = write_definition_file(
        p4, "modbusTCP", "Inverter", "MarsRenewable", "LC-AC", lc_ac_regs, strict_overlap=True
    )
    if not valid:
        raise ValueError(f"Validation failed for {p4.name}: {rep.issues}")
    generated["MarsRenewable_LC-AC.csv"] = p4

    # -------------------------------------------------------------------------
    # 4b. MarsRenewable_LC.csv (Liquid Cooling 1 individually)
    # -------------------------------------------------------------------------
    lc_alone_regs = []
    for item in raw["lc_cooling"]:
        if item["action"] != "0":
            addr = item["addr"]
            tag = "TEMP_AMB" if addr == 30006 else ""
            lc_alone_regs.append(
                {
                    "Info1": "4",
                    "Info2": str(addr),
                    "Info3": item["dtype"],
                    "Name": f"LiquidCooling1 {item['name']}",
                    "Tag": tag,
                    "CoefA": item["coef_a"],
                    "CoefB": 0.0,
                    "Unit": item["unit"],
                    "Action": item["action"],
                }
            )

    p4_lc = out_dir / "MarsRenewable_LC.csv"
    valid, rep = write_definition_file(
        p4_lc, "modbusTCP", "Inverter", "MarsRenewable", "LiquidCooling", lc_alone_regs, strict_overlap=True
    )
    if not valid:
        raise ValueError(f"Validation failed for {p4_lc.name}: {rep.issues}")
    generated["MarsRenewable_LC.csv"] = p4_lc

    # Also generate MarsRenewable_LiquidCooling.csv alias
    p4_lc_alias = out_dir / "MarsRenewable_LiquidCooling.csv"
    write_definition_file(
        p4_lc_alias, "modbusTCP", "Inverter", "MarsRenewable", "LiquidCooling", lc_alone_regs, strict_overlap=True
    )
    generated["MarsRenewable_LiquidCooling.csv"] = p4_lc_alias

    # -------------------------------------------------------------------------
    # 4c. MarsRenewable_AC.csv (Air Conditioning / Air Cooling 1 individually)
    # -------------------------------------------------------------------------
    ac_alone_regs = []
    for item in raw["ac"]:
        if item["action"] != "0":
            addr = item["addr"]
            tag = "TEMP" if addr == 15020 else ("HUMIDITY" if addr == 15027 else "")
            ac_alone_regs.append(
                {
                    "Info1": "4",
                    "Info2": str(addr),
                    "Info3": item["dtype"],
                    "Name": f"AC1 {item['name']}",
                    "Tag": tag,
                    "CoefA": item["coef_a"],
                    "CoefB": 0.0,
                    "Unit": item["unit"],
                    "Action": item["action"],
                }
            )

    p4_ac = out_dir / "MarsRenewable_AC.csv"
    valid, rep = write_definition_file(
        p4_ac, "modbusTCP", "Inverter", "MarsRenewable", "AirCooling", ac_alone_regs, strict_overlap=True
    )
    if not valid:
        raise ValueError(f"Validation failed for {p4_ac.name}: {rep.issues}")
    generated["MarsRenewable_AC.csv"] = p4_ac

    # Also generate MarsRenewable_AirCooling.csv alias
    p4_ac_alias = out_dir / "MarsRenewable_AirCooling.csv"
    write_definition_file(
        p4_ac_alias, "modbusTCP", "Inverter", "MarsRenewable", "AirCooling", ac_alone_regs, strict_overlap=True
    )
    generated["MarsRenewable_AirCooling.csv"] = p4_ac_alias

    # -------------------------------------------------------------------------
    # 5. MarsRenewable_Meter.csv (Device ID 1: Meter 1)
    # -------------------------------------------------------------------------
    meter_regs = []
    meter_tag_map = {
        10000: "E_IMP",
        10010: "E_EXP",
        10020: "P_AC",
    }
    for item in raw["meter"]:
        if item["action"] != "0":
            addr = item["addr"]
            tag = meter_tag_map.get(addr, "")
            meter_regs.append(
                {
                    "Info1": "4",
                    "Info2": str(addr),
                    "Info3": item["dtype"],
                    "Name": f"Meter1 {item['name']}",
                    "Tag": tag,
                    "CoefA": item["coef_a"],
                    "CoefB": 0.0,
                    "Unit": item["unit"],
                    "Action": item["action"],
                }
            )

    p5 = out_dir / "MarsRenewable_Meter.csv"
    valid, rep = write_definition_file(
        p5, "modbusTCP", "Meter", "MarsRenewable", "Meter", meter_regs, strict_overlap=True
    )
    if not valid:
        raise ValueError(f"Validation failed for {p5.name}: {rep.issues}")
    generated["MarsRenewable_Meter.csv"] = p5

    # -------------------------------------------------------------------------
    # 6. MarsRenewable_PV.csv (Device ID 1: PV 1)
    # -------------------------------------------------------------------------
    pv_regs = []
    pv_tag_map = {
        25002: "U_AC_1",
        25004: "U_AC_2",
        25006: "U_AC_3",
        25008: "I_AC_1",
        25010: "I_AC_2",
        25012: "I_AC_3",
        25014: "FREQ",
        25016: "P_AC",
        25018: "Q_AC",
        25020: "S_AC",
        25022: "PF",
        25024: "P_AC_1",
        25026: "P_AC_2",
        25028: "P_AC_3",
        25030: "Q_AC_1",
        25032: "Q_AC_2",
        25034: "Q_AC_3",
        25036: "S_AC_1",
        25038: "S_AC_2",
        25040: "S_AC_3",
        25042: "PF_1",
        25044: "PF_2",
        25046: "PF_3",
        25048: "TEMP_INT",
        25050: "TEMP_AMB",
        25054: "E_TOT",
        25058: "E_DAY",
        25062: "P_DC_MPPT_1",
        25064: "U_DC_MPPT_1",
        25066: "I_DC_MPPT_1",
    }
    used_pv_tags: set[str] = set()

    for item in raw["pv"]:
        if item["action"] != "0":
            addr = item["addr"]
            # Skip the single template row 'Branch 2~n' as we expand branches 1..6 below
            if addr in (25068, 25070, 25072) and "2~n" in item["name"]:
                continue

            tag = pv_tag_map.get(addr, "")
            if tag and tag in used_pv_tags:
                tag = ""
            if tag:
                used_pv_tags.add(tag)

            pv_regs.append(
                {
                    "Info1": "4",
                    "Info2": str(addr),
                    "Info3": item["dtype"],
                    "Name": f"PV1 {item['name']}",
                    "Tag": tag,
                    "CoefA": item["coef_a"],
                    "CoefB": 0.0,
                    "Unit": item["unit"],
                    "Action": item["action"],
                }
            )

    # Expand PV Branches 2 to 6 (stride 6, 25068..25096)
    for b in range(2, 7):
        b_base = 25062 + (b - 1) * 6
        pv_regs.append(
            {
                "Info1": "4",
                "Info2": str(b_base),
                "Info3": "F32",
                "Name": f"PV1 Branch {b} DC Power",
                "Tag": f"P_DC_MPPT_{b}",
                "CoefA": 1.0,
                "CoefB": 0.0,
                "Unit": "kW",
                "Action": "4",
            }
        )
        pv_regs.append(
            {
                "Info1": "4",
                "Info2": str(b_base + 2),
                "Info3": "F32",
                "Name": f"PV1 Branch {b} DC Voltage",
                "Tag": f"U_DC_MPPT_{b}",
                "CoefA": 1.0,
                "CoefB": 0.0,
                "Unit": "V",
                "Action": "4",
            }
        )
        pv_regs.append(
            {
                "Info1": "4",
                "Info2": str(b_base + 4),
                "Info3": "F32",
                "Name": f"PV1 Branch {b} DC Current",
                "Tag": f"I_DC_MPPT_{b}",
                "CoefA": 1.0,
                "CoefB": 0.0,
                "Unit": "A",
                "Action": "4",
            }
        )

    p6 = out_dir / "MarsRenewable_PV.csv"
    valid, rep = write_definition_file(
        p6, "modbusTCP", "Inverter", "MarsRenewable", "PV", pv_regs, strict_overlap=True
    )
    if not valid:
        raise ValueError(f"Validation failed for {p6.name}: {rep.issues}")
    generated["MarsRenewable_PV.csv"] = p6

    return generated


if __name__ == "__main__":
    src_xlsx = Path("stress_test_data") / "MarsLocalController_IO List_ModbusTCP_V7.0.xlsx"
    if not src_xlsx.exists():
        src_xlsx = Path(
            r"\\WDN-FILES.WEBDYN.COM\Support\Equipementiers\MarsRenewable\Sources\MarsLocalController_IO List_ModbusTCP_V7.0.xlsx"
        )

    dest_dir = Path("generated_definitions")
    print(f"Generating Mars Renewable definitions from {src_xlsx} into {dest_dir}...")
    res = generate_mars_requested_definitions(src_xlsx, dest_dir)
    print("Generated successfully:")
    for name, p in res.items():
        print(f"  - {name}: {p} (exists={p.exists()})")
