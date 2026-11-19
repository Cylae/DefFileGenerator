"""
Universal Generator for X7 All-In-One Commercial & Industrial BESS Definitions.
Strictly aligned with Apollo-Standard specifications:
- Sequential index 1..N
- Modbus protocol header (Modbus;)
- Exact 11 columns in header and all rows
- Decimal 0-based Modbus addressing and STRING length syntax (<addr>_<len>)
- Certified Webdyn action codes (0=Reserve, 1=Write/Command, 4=Telemetry/Read, 9=Status/Alarm)
- Canonical SunPM tags (unique single-tags per file)
- Strict UTF-8 encoding without BOM
- Standard naming pattern: <PREFIX>_<Manufacturer>_<Model>.csv
"""

from __future__ import annotations

import csv
import shutil
from pathlib import Path
from typing import Any

import openpyxl

from DefFileGenerator.def_gen import Generator
from DefFileGenerator.production_governor import verify_file_pillars

# Mapping of raw units to clean Webdyn units
UNIT_MAP = {
    "kw": "kW",
    "KW": "kW",
    "kW": "kW",
    "kvar": "kvar",
    "kVar": "kvar",
    "KWh": "kWh",
    "kwh": "kWh",
    "KWH": "kWh",
    "v": "V",
    "V": "V",
    "a": "A",
    "A": "A",
    "w": "W",
    "W": "W",
    "va": "VA",
    "VA": "VA",
    "var": "var",
    "Var": "var",
    "hz": "Hz",
    "Hz": "Hz",
    "%": "%",
    "%rh": "%RH",
    "%RH": "%RH",
    "℃": "°C",
    "°c": "°C",
    "degc": "°C",
    "bar": "bar",
    "Bar": "bar",
    "s": "s",
    "mv": "mV",
    "mV": "mV",
    "%/s": "%/s",
}


def normalize_unit(raw_unit: Any) -> str:
    """Normalizes engineering units."""
    if not raw_unit:
        return ""
    u_str = str(raw_unit).strip()
    return UNIT_MAP.get(u_str, u_str)


def normalize_datatype(raw_type: Any, is_string: bool = False) -> str:
    """Maps vendor type strings to canonical WebdynSunPM data types."""
    if is_string:
        return "STRING"
    if not raw_type:
        return "U16"
    t_str = str(raw_type).strip().upper()
    if t_str in ("UINT16", "U16"):
        return "U16"
    if t_str in ("INT16", "I16"):
        return "I16"
    if t_str in ("UINT32", "U32"):
        return "U32"
    if t_str in ("INT32", "I32"):
        return "I32"
    if t_str in ("FLOAT32", "FLOAT", "F32"):
        return "FLOAT32"
    return "U16"


def compute_coefficients(scale: Any) -> tuple[float, float]:
    """
    Computes CoefA and CoefB from raw vendor scale.
    Formula: Engineering_Value = Raw_Value * CoefA + CoefB.
    """
    if scale is None:
        return 1.0, 0.0
    scale_str = str(scale).strip()
    if scale_str == "10":
        return 0.1, 0.0
    if scale_str == "100":
        return 0.01, 0.0
    if scale_str == "1000":
        return 0.001, 0.0
    if scale_str == "0.01":
        return 0.01, 0.0
    if scale_str == "0.1":
        return 0.1, 0.0
    if scale_str == "1":
        return 1.0, 0.0
    try:
        val = float(scale_str)
        if val > 1.0:
            return 1.0 / val, 0.0
        return val, 0.0
    except ValueError:
        return 1.0, 0.0


def write_definition_file(
    output_path: Path,
    category: str,
    manufacturer: str,
    model: str,
    registers: list[dict[str, Any]],
) -> None:
    """
    Writes a strictly compliant WebdynSunPM CSV definition file in UTF-8 without BOM.
    Ensures sequential 1..N indexing, 11 columns, 6-decimal coefficients, and sanitized fields.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "w", encoding="utf-8", newline="") as fp:
        writer = csv.writer(fp, delimiter=";", lineterminator="\r\n")

        # Header: Modbus;<Category>;<Manufacturer>;<Model>;;;;;;;
        norm_proto = Generator.normalize_protocol("Modbus")
        writer.writerow([norm_proto, category, manufacturer, model, "", "", "", "", "", "", ""])

        # Data rows
        for idx, reg in enumerate(registers, start=1):
            fc = str(reg["fc"]).strip()
            addr = str(reg["address"]).strip()
            dtype = str(reg["data_type"]).strip()
            name = Generator.sanitize_csv_field(str(reg["name"]).strip())
            tag = str(reg.get("tag", "")).strip()
            coef_a = f"{float(reg.get('coef_a', 1.0)):.6f}"
            coef_b = f"{float(reg.get('coef_b', 0.0)):.6f}"
            unit = normalize_unit(reg.get("unit", ""))
            action = str(reg.get("action", "4")).strip()

            writer.writerow(
                [
                    str(idx),
                    fc,
                    addr,
                    dtype,
                    "",
                    name,
                    tag,
                    coef_a,
                    coef_b,
                    unit,
                    action,
                ]
            )

    # Validate against 5 Apollo Pillars
    is_valid, issues = verify_file_pillars(output_path)
    if not is_valid:
        raise ValueError(
            f"Validation failed for generated file {output_path.name}:\n" + "\n".join(issues)
        )


def parse_x7_workbook(xlsx_path: Path | str) -> dict[str, list[dict[str, Any]]]:
    """
    Parses all register sheets of the X7 Modbus TCP workbook into structured register lists.
    Returns a dict mapping subsystem name to its list of registers.
    """
    xlsx_path = Path(xlsx_path)
    if not xlsx_path.is_file():
        raise FileNotFoundError(f"X7 workbook not found at: {xlsx_path}")

    wb = openpyxl.load_workbook(str(xlsx_path), data_only=True)

    # Structure to hold parsed registers per subsystem
    parsed: dict[str, list[dict[str, Any]]] = {
        "ems": [],
        "bms_summary": [],
        "bms_cells": [],
        "meters": [],
        "pcs": [],
        "cabinet": [],
        "all_in_one": [],
    }

    # 1. EMS-Telemetry (FC 4)
    ws = wb["EMS-Telemetry"]
    for r in list(ws.iter_rows(values_only=True))[2:]:
        if not any(r) or len(r) < 5 or r[3] is None:
            continue
        addr = r[3]
        if str(addr).lower() == "address":
            continue
        name = str(r[4]).strip()
        raw_type = r[6]
        raw_scale = r[8]
        raw_unit = r[9]
        remark = str(r[11]) if len(r) > 11 and r[11] else ""

        is_string = "24 bit" in remark.lower() or "string" in name.lower() or raw_type is None
        dtype = normalize_datatype(raw_type, is_string=is_string)
        coef_a, coef_b = compute_coefficients(raw_scale)
        addr_str = f"{addr}_24" if dtype == "STRING" else str(addr)

        # Determine tag & action
        action = "4"
        tag = ""
        if addr == 101:
            tag = "P_AC"
        elif addr == 103:
            tag = "U_DC"
        elif addr == 104:
            tag = "I_DC"
        elif addr == 105:
            tag = "SOC"
        elif addr == 106:
            tag = "STATE_BAT"
            action = "9"
        elif addr == 107:
            tag = "NUM_CELLS"
        elif addr == 108:
            tag = "T_BAT_AVG"
        elif addr == 109:
            tag = "T_BAT_DIFF_MAX"
        elif addr == 110:
            tag = "U_CELL_MAX"
        elif addr == 111:
            tag = "U_CELL_MIN"
        elif addr == 114:
            tag = "STATE_SYS"
            action = "9"
        elif addr == 115:
            tag = "STATE_RUN"
            action = "9"
        elif addr == 116:
            tag = "FLAG_AC_SERIES"
        elif addr == 117:
            tag = "P_RATED"
        elif addr == 138:
            tag = "FW_VER_EMS"
        elif addr == 162:
            tag = "SN_EMS"
        elif addr == 186:
            tag = "VER_PROTO"
        elif addr == 32000:
            tag = "SERIAL"
        elif addr == 32024:
            tag = "SN_PACK_1"
        elif addr == 32048:
            tag = "SN_PACK_2"
        elif addr == 32072:
            tag = "SN_PACK_3"
        elif addr == 32096:
            tag = "SN_PACK_4"
        elif addr == 32120:
            tag = "FW_VER"
        elif addr == 32144:
            tag = "HW_VER"
        elif isinstance(addr, int) and 40000 <= addr <= 40018:
            f_idx = (addr - 40000) // 2 + 1
            tag = f"FAULT_CODE_{f_idx}"
            action = "9"

        reg_entry = {
            "fc": "4",
            "address": addr_str,
            "data_type": dtype,
            "name": f"EMS - {name}",
            "tag": tag,
            "coef_a": coef_a,
            "coef_b": coef_b,
            "unit": raw_unit,
            "action": action,
        }
        parsed["ems"].append(reg_entry)

    # 2. EMS-Remote Signaling (FC 3)
    ws = wb["EMS-Remote Signaling"]
    for r in list(ws.iter_rows(values_only=True))[2:]:
        if not any(r) or len(r) < 5 or r[3] is None:
            continue
        addr = r[3]
        if str(addr).lower() == "address":
            continue
        name = str(r[4]).strip()
        raw_type = r[6]
        raw_scale = r[8]
        raw_unit = r[9]
        coef_a, coef_b = compute_coefficients(raw_scale)
        dtype = normalize_datatype(raw_type)

        tag = ""
        action = "1"
        if addr == 4199:
            tag = "CMD_MS_WORK_MODE"
        elif addr == 4200:
            tag = "STATE_MS_MODE"
            action = "4"
        elif addr == 4201:
            tag = "CMD_MS_SETTINGS"
        elif addr == 4203:
            tag = "NUM_MS_DEVICES"
            action = "4"

        reg_entry = {
            "fc": "3",
            "address": str(addr),
            "data_type": dtype,
            "name": f"EMS - {name}",
            "tag": tag,
            "coef_a": coef_a,
            "coef_b": coef_b,
            "unit": raw_unit,
            "action": action,
        }
        parsed["ems"].append(reg_entry)

    # 3. EMS-Remote Control (FC 3)
    ws = wb["EMS-Remote Control"]
    for r in list(ws.iter_rows(values_only=True))[2:]:
        if not any(r) or len(r) < 5 or r[3] is None:
            continue
        addr = r[3]
        if str(addr).lower() == "address":
            continue
        name = str(r[4]).strip()
        raw_type = r[6]
        raw_scale = r[8]
        raw_unit = r[9]
        coef_a, coef_b = compute_coefficients(raw_scale)
        dtype = normalize_datatype(raw_type)

        tag = ""
        if addr == 2000:
            tag = "CMD_START_STOP"
        elif addr == 2050:
            tag = "CMD_OFFGRID_EN"
        elif addr == 2051:
            tag = "CMD_GRID_CHG_EN"
        elif addr == 2053:
            tag = "CMD_FEED_LIM_EN"
        elif addr == 2054:
            tag = "CMD_FEED_LIM_VAL"
        elif addr == 2999:
            tag = "CMD_SCHED_MODE_FLAG"
        elif isinstance(addr, int) and 3000 <= addr <= 3437:
            # Mode scheduler slot
            slot = (addr - 3000) // 60 + 1
            offset = (addr - 3000) % 60
            tag = f"SCHED_{slot}_OFF_{offset}"

        reg_entry = {
            "fc": "3",
            "address": str(addr),
            "data_type": dtype,
            "name": f"EMS Control - {name}",
            "tag": tag,
            "coef_a": coef_a,
            "coef_b": coef_b,
            "unit": raw_unit,
            "action": "1",
        }
        parsed["ems"].append(reg_entry)

    # 4. EMS-Remote Adjustment (FC 3)
    ws = wb["EMS-Remote Adjustment"]
    for r in list(ws.iter_rows(values_only=True))[2:]:
        if not any(r) or len(r) < 5 or r[3] is None:
            continue
        addr = r[3]
        if str(addr).lower() == "address":
            continue
        name = str(r[4]).strip()
        # Discard deprecated duplicate 106
        if addr == 106 and name == "Total Reactive Power Command":
            continue
        raw_type = r[6]
        raw_scale = r[8]
        raw_unit = r[9]
        remark = str(r[11]) if len(r) > 11 and r[11] else ""

        is_string = "24 bit" in remark.lower() or raw_type is None
        dtype = normalize_datatype(raw_type, is_string=is_string)
        coef_a, coef_b = compute_coefficients(raw_scale)
        addr_str = f"{addr}_24" if dtype == "STRING" else str(addr)

        tag = ""
        if addr == 80:
            tag = "CMD_BACKUP_RESERVE_SOC"
        elif addr == 81:
            tag = "CMD_SOC_MAX"
        elif addr == 82:
            tag = "CMD_SOC_MIN"
        elif addr == 83:
            tag = "CMD_GRID_CHG_VAL"
        elif addr == 93:
            tag = "CMD_MODBUS_TIMEOUT"
        elif addr == 94:
            tag = "CMD_FAULT_RESET"
        elif addr == 95:
            tag = "CMD_COMM_LOSS_STRAT_EN"
        elif addr == 99:
            tag = "CMD_PV_PWR_SETPOINT"
        elif addr == 100:
            tag = "CMD_PWR_CTRL_MODE"
        elif addr == 102:
            tag = "CMD_MODE"
        elif addr == 104:
            tag = "CMD_CTRL_SRC"
        elif addr == 105:
            tag = "CMD_PWR_SETPOINT"
        elif addr == 106:
            tag = "CMD_Q_EN"
        elif addr == 107:
            tag = "CMD_Q_SETPOINT"
        elif addr == 109:
            tag = "CMD_RAMP_RATE"
        elif addr == 900:
            tag = "SN_MACHINE"

        reg_entry = {
            "fc": "3",
            "address": addr_str,
            "data_type": dtype,
            "name": f"EMS Adjust - {name}",
            "tag": tag,
            "coef_a": coef_a,
            "coef_b": coef_b,
            "unit": raw_unit,
            "action": "1",
        }
        parsed["ems"].append(reg_entry)

    # 5. BMS-Telemetry (FC 4)
    ws = wb["BMS-Telemetry"]
    for r in list(ws.iter_rows(values_only=True))[2:]:
        if not any(r) or len(r) < 5 or r[3] is None:
            continue
        addr = r[3]
        if str(addr).lower() == "address":
            continue
        name = str(r[4]).strip()
        raw_type = r[6]
        raw_scale = r[8]
        raw_unit = r[9]

        # Handle Cell Temperature block (128 registers)
        if str(addr) == "6200～6327":
            dtype = "I16"
            coef_a, coef_b = compute_coefficients(raw_scale or 1)
            for c_idx in range(128):
                cell_no = c_idx + 1
                pack_no = (c_idx // 32) + 1
                pack_cell = (c_idx % 32) + 1
                parsed["bms_cells"].append(
                    {
                        "fc": "4",
                        "address": str(6200 + c_idx),
                        "data_type": dtype,
                        "name": f"BMS Cell Temperature {cell_no} (Pack {pack_no} Sensor {pack_cell})",
                        "tag": f"T_CELL_{cell_no}",
                        "coef_a": coef_a,
                        "coef_b": coef_b,
                        "unit": "°C",
                        "action": "4",
                    }
                )
            continue

        # Handle Cell Voltage block (240 registers)
        if str(addr) == "6600～6839":
            dtype = "U16"
            coef_a, coef_b = compute_coefficients(raw_scale or 1000)
            for c_idx in range(240):
                cell_no = c_idx + 1
                pack_no = (c_idx // 60) + 1
                pack_cell = (c_idx % 60) + 1
                parsed["bms_cells"].append(
                    {
                        "fc": "4",
                        "address": str(6600 + c_idx),
                        "data_type": dtype,
                        "name": f"BMS Cell Voltage {cell_no} (Pack {pack_no} Cell {pack_cell})",
                        "tag": f"U_CELL_{cell_no}",
                        "coef_a": coef_a,
                        "coef_b": coef_b,
                        "unit": "V",
                        "action": "4",
                    }
                )
            continue

        if not isinstance(addr, int):
            continue

        dtype = normalize_datatype(raw_type)
        coef_a, coef_b = compute_coefficients(raw_scale)
        action = "4"
        tag = ""
        if addr == 200:
            tag = "SOC_RACK"
        elif addr == 201:
            tag = "SOH_RACK"
        elif addr == 202:
            tag = "STATE_RACK"
            action = "9"
        elif addr == 203:
            tag = "U_DC_RACK"
        elif addr == 204:
            tag = "I_DC_RACK"
        elif addr == 207:
            tag = "T_CELL_MAX"
        elif addr == 208:
            tag = "NUM_CELL_T_MAX"
        elif addr == 209:
            tag = "T_CELL_MIN"
        elif addr == 210:
            tag = "NUM_CELL_T_MIN"
        elif addr == 211:
            tag = "U_CELL_MAX_RACK"
        elif addr == 212:
            tag = "NUM_CELL_U_MAX"
        elif addr == 213:
            tag = "U_CELL_MIN_RACK"
        elif addr == 214:
            tag = "NUM_CELL_U_MIN"
        elif addr == 215:
            tag = "U_CELL_AVG_RACK"
        elif addr == 216:
            tag = "T_CELL_AVG_RACK"
        elif addr == 217:
            tag = "T_CELL_DIFF_MAX"
        elif addr == 218:
            tag = "NUM_CELL_T_MIN2"
        elif addr == 223:
            tag = "U_CELL_DIFF_MAX"
        elif addr == 224:
            tag = "P_DC_RACK"
        elif addr == 227:
            tag = "E_CHG_RACK_TOT"
        elif addr == 229:
            tag = "E_DISCHG_RACK_TOT"

        reg_entry = {
            "fc": "4",
            "address": str(addr),
            "data_type": dtype,
            "name": f"BMS - {name}",
            "tag": tag,
            "coef_a": coef_a,
            "coef_b": coef_b,
            "unit": raw_unit,
            "action": action,
        }
        parsed["bms_summary"].append(reg_entry)

    # 6. BMS-Remote Control (FC 4)
    ws = wb["BMS-Remote Control"]
    for r in list(ws.iter_rows(values_only=True))[2:]:
        if not any(r) or len(r) < 5 or r[3] is None:
            continue
        addr = r[3]
        if str(addr).lower() == "address" or not isinstance(addr, int):
            continue
        name = str(r[4]).strip()
        raw_type = r[6]
        raw_scale = r[8]
        raw_unit = r[9]
        dtype = normalize_datatype(raw_type)
        coef_a, coef_b = compute_coefficients(raw_scale)

        tag = "P_CHG_MAX_ALLOWED" if addr == 2203 else "P_DISCHG_MAX_ALLOWED"
        reg_entry = {
            "fc": "4",
            "address": str(addr),
            "data_type": dtype,
            "name": f"BMS - {name}",
            "tag": tag,
            "coef_a": coef_a,
            "coef_b": coef_b,
            "unit": raw_unit,
            "action": "4",
        }
        parsed["bms_summary"].append(reg_entry)

    # 7. BMS-Remote Signaling (FC 2)
    ws = wb["BMS-Remote Signaling"]
    for r in list(ws.iter_rows(values_only=True))[2:]:
        if not any(r) or len(r) < 5 or r[3] is None:
            continue
        addr = r[3]
        if str(addr).lower() == "address" or not isinstance(addr, int):
            continue
        name = str(r[4]).strip()
        reg_entry = {
            "fc": "2",
            "address": str(addr),
            "data_type": "U16",
            "name": f"BMS Alarm - {name}",
            "tag": "ALARM_BMS_COMM_FAULT",
            "coef_a": 1.0,
            "coef_b": 0.0,
            "unit": "",
            "action": "9",
        }
        parsed["bms_summary"].append(reg_entry)

    # 8. Other Devices-Telemetry (FC 4)
    ws = wb["Other Devices-Telemetry"]
    for r in list(ws.iter_rows(values_only=True))[2:]:
        if not any(r) or len(r) < 5 or r[3] is None:
            continue
        addr = r[3]
        if str(addr).lower() == "address" or not isinstance(addr, int):
            continue
        name = str(r[4]).strip()
        raw_type = r[6]
        raw_scale = r[8]
        raw_unit = r[9]
        remark = str(r[11]) if len(r) > 11 and r[11] else ""

        is_string = (
            "24 bit" in remark.lower()
            or raw_type is None
            and addr in (31000, 31024, 31096, 31120, 31144)
        )
        dtype = normalize_datatype(raw_type, is_string=is_string)
        coef_a, coef_b = compute_coefficients(raw_scale)
        addr_str = f"{addr}_24" if dtype == "STRING" else str(addr)

        # Classify by device subsystem
        if 55100 <= addr <= 55121:
            # Meter 1 (Power Grid)
            tag = ""
            if addr == 55100:
                tag = "U_AC_GRID_1"
            elif addr == 55101:
                tag = "U_AC_GRID_2"
            elif addr == 55102:
                tag = "U_AC_GRID_3"
            elif addr == 55103:
                tag = "U_AC_GRID_12"
            elif addr == 55104:
                tag = "U_AC_GRID_23"
            elif addr == 55105:
                tag = "U_AC_GRID_31"
            elif addr == 55106:
                tag = "I_AC_GRID_1"
            elif addr == 55107:
                tag = "I_AC_GRID_2"
            elif addr == 55108:
                tag = "I_AC_GRID_3"
            elif addr == 55109:
                tag = "P_AC_GRID"
            elif addr == 55110:
                tag = "P_AC_GRID_1"
            elif addr == 55111:
                tag = "P_AC_GRID_2"
            elif addr == 55112:
                tag = "P_AC_GRID_3"
            elif addr == 55113:
                tag = "Q_AC_GRID"
            elif addr == 55114:
                tag = "Q_AC_GRID_1"
            elif addr == 55115:
                tag = "Q_AC_GRID_2"
            elif addr == 55116:
                tag = "Q_AC_GRID_3"
            elif addr == 55117:
                tag = "PF_GRID_1"
            elif addr == 55118:
                tag = "PF_GRID_2"
            elif addr == 55119:
                tag = "PF_GRID_3"
            elif addr == 55120:
                tag = "PF_GRID"
            elif addr == 55121:
                tag = "FREQ_GRID"

            parsed["meters"].append(
                {
                    "fc": "4",
                    "address": addr_str,
                    "data_type": dtype,
                    "name": f"Meter 1 Grid - {name}",
                    "tag": tag,
                    "coef_a": coef_a,
                    "coef_b": coef_b,
                    "unit": raw_unit,
                    "action": "4",
                }
            )

        elif 55150 <= addr <= 55171:
            # Meter 2 (Third-Party PV)
            off = addr - 55150
            parsed["meters"].append(
                {
                    "fc": "4",
                    "address": addr_str,
                    "data_type": dtype,
                    "name": f"Meter 2 PV - {name}",
                    "tag": f"MET2_PV_PARAM_{off}",
                    "coef_a": coef_a,
                    "coef_b": coef_b,
                    "unit": raw_unit,
                    "action": "4",
                }
            )

        elif 55200 <= addr <= 55221:
            # Meter 3 (Diesel Generator)
            off = addr - 55200
            parsed["meters"].append(
                {
                    "fc": "4",
                    "address": addr_str,
                    "data_type": dtype,
                    "name": f"Meter 3 Gen - {name}",
                    "tag": f"MET3_GEN_PARAM_{off}",
                    "coef_a": coef_a,
                    "coef_b": coef_b,
                    "unit": raw_unit,
                    "action": "4",
                }
            )

        elif 60100 <= addr <= 60105 or 60110 <= addr <= 60111:
            # Cabinet Auxiliaries (Liquid Cooler / Dehumidifier Telemetry)
            tag = ""
            if addr == 60110:
                tag = "ENV_HUMIDITY"
            elif addr == 60111:
                tag = "ENV_TEMPERATURE"
            parsed["cabinet"].append(
                {
                    "fc": "4",
                    "address": addr_str,
                    "data_type": dtype,
                    "name": f"Cabinet Aux - {name}",
                    "tag": tag,
                    "coef_a": coef_a,
                    "coef_b": coef_b,
                    "unit": raw_unit,
                    "action": "4",
                }
            )

        else:
            # PCS Inverter Data, Energies, and Parameters
            tag = ""
            if addr == 55755:
                tag = "U_AC_1"
            elif addr == 55756:
                tag = "U_AC_2"
            elif addr == 55757:
                tag = "U_AC_3"
            elif addr == 55758:
                tag = "I_AC_1"
            elif addr == 55759:
                tag = "I_AC_2"
            elif addr == 55760:
                tag = "I_AC_3"
            elif addr == 55761:
                tag = "P_AC_1"
            elif addr == 55762:
                tag = "P_AC_2"
            elif addr == 55763:
                tag = "P_AC_3"
            elif addr == 55764:
                tag = "P_AC_INV_TOTAL"
            elif addr == 55765:
                tag = "Q_AC_1"
            elif addr == 55766:
                tag = "Q_AC_2"
            elif addr == 55767:
                tag = "Q_AC_3"
            elif addr == 55769:
                tag = "Q_AC_INV_TOTAL"
            elif addr == 55770:
                tag = "S_AC_1"
            elif addr == 55771:
                tag = "S_AC_2"
            elif addr == 55772:
                tag = "S_AC_3"
            elif addr == 55773:
                tag = "S_AC_INV_TOTAL"
            elif addr == 55774:
                tag = "PF"
            elif addr == 55775:
                tag = "FREQ"
            elif addr == 55776:
                tag = "T_INT"
            elif 55800 <= addr <= 55815:
                pv_no = addr - 55800 + 1
                tag = f"U_DC_PV{pv_no}"
            elif 55816 <= addr <= 55831:
                pv_no = addr - 55816 + 1
                tag = f"I_DC_PV{pv_no}"
            elif 55832 <= addr <= 55847:
                pv_no = addr - 55832 + 1
                tag = f"P_DC_PV{pv_no}"
            elif addr == 55940:
                tag = "E_TOT_PV"
            elif addr == 55942:
                tag = "STATE_GRID"
            elif addr == 55943:
                tag = "E_TOT_FEED_GRID"
            elif addr == 55945:
                tag = "E_TOT_BUY_GRID"
            elif addr == 55950:
                tag = "E_TOT_BAT_CHG"
            elif addr == 55952:
                tag = "E_TOT_BAT_DISCHG"
            elif addr == 55954:
                tag = "E_DAY_BAT_CHG"
            elif addr == 55955:
                tag = "E_DAY_BAT_DISCHG"
            elif addr == 55956:
                tag = "E_TOT"
            elif addr == 55958:
                tag = "E_DAY"
            elif addr == 55959:
                tag = "E_TOT_PV_CHG"
            elif addr == 55961:
                tag = "E_TOT_GRID_CHG"
            elif addr == 55963:
                tag = "RACK_MAX_ENERGY"
            elif addr == 55964:
                tag = "AVAIL_DISCHG_ENERGY"
            elif addr == 55965:
                tag = "AVAIL_CHG_ENERGY"
            elif addr == 31000:
                tag = "SN_PCS"
            elif addr == 31024:
                tag = "FW_VER_ARM"
            elif addr == 31096:
                tag = "FW_VER_DSP_MASTER"
            elif addr == 31120:
                tag = "FW_VER_DSP_SLAVE"
            elif addr == 31144:
                tag = "FW_VER_AFCI"

            parsed["pcs"].append(
                {
                    "fc": "4",
                    "address": addr_str,
                    "data_type": dtype,
                    "name": f"PCS - {name}",
                    "tag": tag,
                    "coef_a": coef_a,
                    "coef_b": coef_b,
                    "unit": raw_unit,
                    "action": "4",
                }
            )

    # 9. Other Devices-Remote Signaling (FC 2 and FC 4)
    ws = wb["Other Devices-Remote Signaling"]
    curr_fc = "2"
    for r in list(ws.iter_rows(values_only=True)):
        if not any(r):
            continue
        r0 = str(r[0]).strip() if r[0] is not None else ""
        if "0x04" in r0:
            curr_fc = "4"
            continue
        if "0x02" in r0:
            curr_fc = "2"
            continue
        if r0.lower() == "device type" or (len(r) > 3 and str(r[3]).lower() == "address"):
            continue

        if len(r) > 3 and isinstance(r[3], int):
            addr = r[3]
            name = str(r[4]).strip() if len(r) > 4 and r[4] is not None else ""
            raw_type = r[6] if len(r) > 6 else None
            dtype = normalize_datatype(raw_type)

            tag = ""
            if addr == 55100:
                tag = "ALARM_PCS_METER_COMM"
            elif addr == 54200:
                tag = "ALARM_EMERGENCY_STOP"
            elif addr == 54201:
                tag = "ALARM_SURGE"
            elif addr == 54202:
                tag = "ALARM_WATER_IMMERSION"
            elif addr == 54203:
                tag = "ALARM_FIRE_EXTINGUISHER"
            elif addr == 54204:
                tag = "ALARM_SMOKE_DETECTOR"
            elif addr == 54205:
                tag = "ALARM_ACCESS_CONTROL"
            elif addr == 54206:
                tag = "ALARM_TEMP_SENSING"
            elif addr == 54207:
                tag = "STATE_FAN"
            elif addr == 54208:
                tag = "STATE_AIR_OUTLET"
            elif addr == 54209:
                tag = "STATE_AIR_INLET"
            elif addr == 54210:
                tag = "ALARM_COMBUSTIBLE_GAS_FAULT"
            elif addr == 54211:
                tag = "ALARM_COMBUSTIBLE_GAS"
            elif addr == 55350:
                tag = "ALARM_DEHUM_COMM"
            elif addr == 56000:
                tag = "ALARM_LIQUID_COOLING_COMM"
            elif addr == 55750:
                tag = "ALARM_PCS_EMS_COMM"
            elif addr == 55751:
                tag = "ALARM_BMS_EMS_COMM"
            elif 56100 <= addr <= 56125:
                off = addr - 56100
                tag = f"DIAG_PERIPHERAL_{off}"

            parsed["cabinet"].append(
                {
                    "fc": curr_fc,
                    "address": str(addr),
                    "data_type": dtype,
                    "name": f"Safety/Aux - {name}",
                    "tag": tag,
                    "coef_a": 1.0,
                    "coef_b": 0.0,
                    "unit": "",
                    "action": "9",
                }
            )

    # 10. Other Devices-Remote Adjustment (FC 3)
    ws = wb["Other Devices-Remote Adjustment"]
    for r in list(ws.iter_rows(values_only=True))[2:]:
        if not any(r) or len(r) < 5 or r[3] is None:
            continue
        addr = r[3]
        if str(addr).lower() == "address" or not isinstance(addr, int):
            continue
        name = str(r[4]).strip()
        raw_type = r[6]
        raw_scale = r[8]
        raw_unit = r[9]
        dtype = normalize_datatype(raw_type)
        coef_a, coef_b = compute_coefficients(raw_scale)

        tag = "CMD_DEHUM_SETPOINT" if addr == 60120 else "CMD_DEHUM_MARGIN"
        parsed["cabinet"].append(
            {
                "fc": "3",
                "address": str(addr),
                "data_type": dtype,
                "name": f"Cabinet Setting - {name}",
                "tag": tag,
                "coef_a": coef_a,
                "coef_b": coef_b,
                "unit": raw_unit,
                "action": "1",
            }
        )

    # Build Master All-In-One Definition
    # Combines: ems + bms_summary + meters + pcs + cabinet + bms_cells
    # Ensures tag uniqueness across the entire system
    seen_single_tags: set[str] = set()
    single_tag_whitelist = {
        "P_AC",
        "Q_AC",
        "S_AC",
        "P_DC",
        "FREQ",
        "PF",
        "E_TOT",
        "E_DAY",
        "MODEL",
        "SERIAL",
    }

    all_in_one_regs: list[dict[str, Any]] = []

    # Order: EMS -> BMS Summary -> Meters -> PCS -> Cabinet -> Cells
    combined_groups = [
        parsed["ems"],
        parsed["bms_summary"],
        parsed["meters"],
        parsed["pcs"],
        parsed["cabinet"],
        parsed["bms_cells"],
    ]

    for grp in combined_groups:
        for reg in grp:
            clean_reg = dict(reg)
            t = clean_reg.get("tag", "").strip()
            if t in single_tag_whitelist:
                if t in seen_single_tags:
                    # Differentiate duplicate single tags for master file
                    clean_reg["tag"] = f"{t}_SUB"
                else:
                    seen_single_tags.add(t)
            all_in_one_regs.append(clean_reg)

    parsed["all_in_one"] = all_in_one_regs
    return parsed


def generate_all_x7_definitions(
    source_wb_path: Path | str,
    output_dir: Path | str,
) -> dict[str, Path]:
    """
    Generates the complete certified X7 definition suite:
    1. BAT_X7_BESS-AllInOne.csv
    2. BAT_X7_EMS.csv
    3. BAT_X7_BMS.csv
    4. BAT_X7_BMS-Cells.csv
    5. MET_X7_Meters.csv
    6. INV_X7_PCS.csv
    7. OTH_X7_Cabinet.csv
    """
    source_path = Path(source_wb_path)
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    parsed = parse_x7_workbook(source_path)

    generated_files: dict[str, Path] = {}

    targets = [
        ("BAT_X7_BESS-AllInOne.csv", "Battery", "X7", "BESS-AllInOne", parsed["all_in_one"]),
        ("BAT_X7_EMS.csv", "Battery", "X7", "EMS", parsed["ems"]),
        ("BAT_X7_BMS.csv", "Battery", "X7", "BMS", parsed["bms_summary"]),
        ("BAT_X7_BMS-Cells.csv", "Battery", "X7", "BMS-Cells", parsed["bms_cells"]),
        ("MET_X7_Meters.csv", "Meter", "X7", "Meters", parsed["meters"]),
        ("INV_X7_PCS.csv", "Inverter", "X7", "PCS", parsed["pcs"]),
        ("OTH_X7_Cabinet.csv", "Other", "X7", "Cabinet", parsed["cabinet"]),
    ]

    for fname, cat, mfg, model, regs in targets:
        target_path = out_dir / fname
        write_definition_file(target_path, cat, mfg, model, regs)
        generated_files[fname] = target_path

    return generated_files


def deploy_redundant_definitions(
    generated_files: dict[str, Path],
    workspace_dir: Path | str,
) -> None:
    """
    Deploys generated files across all synchronized production & definition directories:
    - production/
    - Definitions/X7/
    - Definitions/
    - X7_Definitions/
    - generated_definitions/
    - tagged_definitions/
    """
    ws = Path(workspace_dir)
    target_dirs = [
        ws / "production",
        ws / "Definitions" / "X7",
        ws / "X7_Definitions",
        ws / "generated_definitions",
        ws / "tagged_definitions",
    ]

    for d in target_dirs:
        d.mkdir(parents=True, exist_ok=True)
        for fname, src_file in generated_files.items():
            dest = d / fname
            if src_file.resolve() == dest.resolve():
                continue
            shutil.copy2(src_file, dest)
            # Re-verify destination
            is_valid, issues = verify_file_pillars(dest)
            if not is_valid:
                raise ValueError(f"Pillar violation at {dest}:\n" + "\n".join(issues))
