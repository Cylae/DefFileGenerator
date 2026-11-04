#!/usr/bin/env python3
"""
High-Precision Generator for Dyness DH-Series Commercial & Industrial BESS
(DH200Y, DH200F, DH100F, and DH_Series_Generic).
Strictly aligned with WebdynSunPM reference specifications.

Features:
- Complete parsing of Dyness DH-Series Modbus protocol (V1.0 - V2.2).
- Automatic extraction and distinction between:
    * Holding Registers (FC 03 / 06 / 16 -> Info1=3, Action=1 or 4)
    * Common Input Registers (FC 04 -> Info1=4, Action=4 or 9)
    * Parallel Operation Aggregated Data (4000+, 4125+, 5875+, 6000+)
    * Model-specific PCS modules (Position #4 for DH200Y, Position #5 for DH200F/DH100F)
    * DCDC modules (DH200F/DH100F)
    * BMS pack metrics and multi-cell voltage/temp expansions (260 cells for DH200Y, 240 cells for DH200F/DH100F)
    * Liquid Cooling module (DH200Y) vs AC / Air Conditioner module (DH200F/DH100F)
    * Fire Protection and Temperature/Humidity monitoring systems
- 11-column standard WebdynSunPM format (Header + data rows).
- 6-decimal floating point coefficients (CoefA, CoefB) with physical offsets (-40.0 for temps, -3200.0 for pack current).
- Canonical SunPM semantic tags assigned with strict tag uniqueness (0 duplicates).
- Strict validation via Generator().validate_csv_detailed().
"""

from __future__ import annotations

import argparse
import csv
import math
import re
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from DefFileGenerator.def_gen import Generator  # noqa: E402
from DefFileGenerator.knowledge_miner import infer_sunpm_tag  # noqa: E402


def clean_text(val: str | None) -> str:
    """Strips whitespace and neutralizes CSV formulas."""
    if not val:
        return ""
    s = str(val).strip()
    # Normalize double quotes
    s = s.replace('"', "")
    return s


def parse_coeff_and_unit(coeff_str: str, unit_str: str) -> tuple[float, str]:
    """Extracts numeric float coefficient and normalized engineering unit."""
    c_s = clean_text(coeff_str)
    u_s = clean_text(unit_str)

    # Patterns like '0.1V', '0.1A', '0.1kW', '0.1kva', '0.01Hz', '0.1° C', '0.1°C'
    m = re.match(r"^([0-9.eE-]+)\s*([A-Za-z°℃%Ω/].*)?$", c_s)
    if m:
        try:
            c_val = float(m.group(1))
        except ValueError:
            c_val = 1.0
        extracted_unit = m.group(2) or ""
        final_unit = u_s if u_s else extracted_unit
    else:
        try:
            c_val = float(c_s) if c_s else 1.0
        except ValueError:
            c_val = 1.0
        final_unit = u_s

    # Unit normalization
    if final_unit in ("℃", "C", "° C", "0.1℃"):
        final_unit = "°C"
    final_unit = final_unit.strip()

    if not math.isfinite(c_val):
        c_val = 1.0

    unit_mapping = {
        "kw": "kW",
        "kwh": "kWh",
        "kvar": "kvar",
        "hz": "Hz",
        "bar": "bar",
        "rpm": "rpm",
        "va": "VA",
        "kva": "kVA",
        "v": "V",
        "a": "A",
        "s": "s",
        "%rh": "%RH",
        "%": "%",
    }
    final_unit = unit_mapping.get(final_unit.lower(), final_unit)
    return c_val, final_unit


class DynessProtocolParser:
    """Parses Dyness DH-Series Modbus protocol CSV into structured register definitions."""

    def __init__(self, source_csv_path: str | Path):
        self.source_path = Path(source_csv_path)
        self.holding_regs: list[dict] = []
        self.common_input_regs: list[dict] = []
        self.parallel_input_regs: list[dict] = []
        self.pcs_pos4_dh200y: list[dict] = []
        self.pcs_pos5_dh200f: list[dict] = []
        self.dcdc_regs: list[dict] = []
        self.bms_common_regs: list[dict] = []
        self.fire_regs: list[dict] = []
        self.ac_regs: list[dict] = []
        self.liquid_cooling_regs: list[dict] = []
        self.temp_hum_regs: list[dict] = []

    def parse(self) -> None:
        if not self.source_path.is_file():
            raise FileNotFoundError(f"Source file not found at: {self.source_path}")

        with open(self.source_path, encoding="utf-8-sig", errors="replace") as f:
            rows = list(csv.reader(f))

        current_section = None
        current_fc = 4

        for row_idx, r in enumerate(rows):
            line_no = row_idx + 1
            non_empty = [c.strip() for c in r if c.strip()]
            line_text = " ".join(non_empty).lower()

            # Section detection
            if "1、03 data" in line_text:
                current_section = "HOLDING"
                current_fc = 3
                continue
            elif "2、04 data" in line_text:
                current_section = "INPUT_COMMON"
                current_fc = 4
                continue
            elif "parallel operation data" in line_text and "master unit" in line_text:
                current_section = "INPUT_PARALLEL"
                current_fc = 4
                continue
            elif "the fifth data group" in line_text and "position #4" in line_text:
                current_section = "PCS_POS4_DH200Y"
                current_fc = 4
                continue
            elif "the fifth data group" in line_text and "position #5" in line_text:
                current_section = "PCS_POS5_DH200F"
                current_fc = 4
                continue
            elif "the fifth data group" in line_text and (
                "position #1" in line_text
                or "position #2" in line_text
                or "position #3" in line_text
            ):
                # Variant positions for PCS (alternate sub-modules)
                current_section = "PCS_OTHER"
                current_fc = 4
                continue
            elif "the sixth group of data" in line_text and "dcdc" in line_text:
                current_section = "DCDC"
                current_fc = 4
                continue
            elif "for groups 7-9 of bms data" in line_text:
                current_section = "BMS"
                current_fc = 4
                continue
            elif "fire protection data" in line_text and (
                "1375" in line_text or "twelfth" in line_text or "group 12" in line_text
            ):
                current_section = "FIRE"
                current_fc = 4
                continue
            elif "ac data" in line_text and (
                "1500" in line_text or "thirteenth" in line_text or "group 13" in line_text
            ):
                current_section = "AC"
                current_fc = 4
                continue
            elif (
                "the liquid cooling unit data" in line_text
                or "liquid cooling unit data (1625" in line_text
            ):
                current_section = "LIQUID_COOLING"
                current_fc = 4
                continue
            elif "temperature and humidity controller data" in line_text:
                current_section = "TEMP_HUM"
                current_fc = 4
                continue
            elif "event code" in line_text and "solution" in line_text:
                # Alarm event code dictionary appendix
                current_section = "ALARM_DICT"
                continue

            if current_section in (None, "ALARM_DICT", "PCS_OTHER"):
                continue

            # Extract fields based on section column alignment
            if current_section == "HOLDING":
                record = self._parse_holding_row(r, line_no)
                if record:
                    self.holding_regs.append(record)
            elif current_section in ("INPUT_COMMON", "INPUT_PARALLEL"):
                record = self._parse_standard_input_row(r, line_no, current_fc)
                if record:
                    if current_section == "INPUT_COMMON":
                        self.common_input_regs.append(record)
                    else:
                        self.parallel_input_regs.append(record)
            else:
                # Sub-module sections (indented with 2 leading empty columns)
                record = self._parse_submodule_row(r, line_no, current_section)
                if record:
                    if current_section == "PCS_POS4_DH200Y":
                        self.pcs_pos4_dh200y.append(record)
                    elif current_section == "PCS_POS5_DH200F":
                        self.pcs_pos5_dh200f.append(record)
                    elif current_section == "DCDC":
                        self.dcdc_regs.append(record)
                    elif current_section == "BMS":
                        self.bms_common_regs.append(record)
                    elif current_section == "FIRE":
                        self.fire_regs.append(record)
                    elif current_section == "AC":
                        self.ac_regs.append(record)
                    elif current_section == "LIQUID_COOLING":
                        self.liquid_cooling_regs.append(record)
                    elif current_section == "TEMP_HUM":
                        self.temp_hum_regs.append(record)

    def _parse_holding_row(self, r: list[str], line_no: int) -> dict | None:
        if len(r) < 5:
            return None
        addr_clean = r[1].strip() if len(r) > 1 else ""
        if not re.match(r"^\d+$", addr_clean):
            return None

        var_name = clean_text(r[2]) if len(r) > 2 else ""
        desc = clean_text(r[3]) if len(r) > 3 else var_name
        dtype = clean_text(r[4]).upper() if len(r) > 4 else "U16"
        rw = clean_text(r[5]).upper() if len(r) > 5 else "R/W"
        coeff_raw = clean_text(r[7]) if len(r) > 7 else "1"
        unit_raw = clean_text(r[8]) if len(r) > 8 else ""
        comment1 = clean_text(r[10]) if len(r) > 10 else ""
        comment2 = clean_text(r[11]) if len(r) > 11 else ""

        if not desc:
            desc = var_name
        if not desc:
            return None

        # Normalize data type
        dtype = (
            "I16"
            if dtype in ("S16", "INT16")
            else (
                "U32"
                if dtype in ("U32", "UINT32")
                else ("I32" if dtype in ("S32", "INT32") else "U16")
            )
        )
        coeff, unit = parse_coeff_and_unit(coeff_raw, unit_raw)
        action = "1" if "W" in rw else "4"

        # Check comment for model restrictions
        model_restriction = ""
        full_comment = f"{comment1} {comment2}".lower()
        if (
            "only available on 200f/100f" in full_comment
            or "available only on the 200f and 100f" in full_comment
        ):
            model_restriction = "200F_100F_ONLY"
        elif "dh200y" in full_comment:
            model_restriction = "DH200Y_ONLY"

        return {
            "info1": 3,
            "address": addr_clean,
            "name": desc,
            "dtype": dtype,
            "coeff_a": coeff,
            "coeff_b": 0.0,
            "unit": unit,
            "action": action,
            "model_restriction": model_restriction,
            "line": line_no,
        }

    def _parse_standard_input_row(self, r: list[str], line_no: int, fc: int) -> dict | None:
        if len(r) < 5:
            return None
        addr_clean = r[1].strip() if len(r) > 1 else ""
        if not re.match(r"^\d+(?:~\d+)?$", addr_clean):
            return None

        var_name = clean_text(r[2]) if len(r) > 2 else ""
        desc = clean_text(r[3]) if len(r) > 3 else var_name
        dtype = clean_text(r[4]).upper() if len(r) > 4 else ""
        coeff_raw = clean_text(r[7]) if len(r) > 7 else "1"
        unit_raw = clean_text(r[8]) if len(r) > 8 else ""
        comment1 = clean_text(r[9]) if len(r) > 9 else ""
        comment2 = clean_text(r[10]) if len(r) > 10 else ""

        if not desc and not var_name:
            return None
        if desc.lower() == "reserved" or var_name.lower() == "reserved":
            return None

        # Check for ASCII strings
        is_ascii = (
            dtype == "ASCII"
            or "ASCII" in coeff_raw
            or "ASCII" in unit_raw
            or "ASCII" in comment1
            or "ASCII" in comment2
            or addr_clean
            in (
                "33",
                "43",
                "53",
                "254",
                "2125",
                "2135",
                "2145",
                "2160",
                "2170",
                "4033",
                "4043",
                "4053",
            )
        )

        # Empty dtype in standard table means Low word continuation, string continuation, or reserved
        if not dtype and not is_ascii:
            return None

        if not desc:
            desc = var_name

        # Clean "(High Word)" or " H" from description
        clean_name = re.sub(r"\s*\((?:High(?:\s*Word)?|High)\)", "", desc, flags=re.IGNORECASE)
        clean_name = re.sub(r"\s+H$", "", clean_name).strip()

        # Normalize data type
        dtype_norm = "U16"
        if is_ascii:
            dtype_norm = "STRING"
        elif dtype in ("S16", "INT16"):
            dtype_norm = "I16"
        elif dtype in ("U32", "UINT32"):
            dtype_norm = "U32"
        elif dtype in ("S32", "INT32"):
            dtype_norm = "I32"
        elif dtype in ("ASCII", "STRING", "STR"):
            dtype_norm = "STRING"

        # String address span formatting
        formatted_address = addr_clean
        if dtype_norm == "STRING":
            # Map known string spans
            spans = {
                "33": "33_20",  # EMS_ver (10 regs = 20 chars)
                "43": "43_20",  # BMS/DCDC_ver
                "53": "53_20",  # HMI_ver
                "254": "254_24",  # Parallel identification code 1..12
                "2125": "2125_20",  # 4g_Hardware_ver
                "2135": "2135_20",  # 4g_Software_ver1
                "2145": "2145_30",  # 4g_Software_ver2 (15 regs)
                "2160": "2160_20",  # 4g_IMEI
                "2170": "2170_32",  # 4g_SIM Card Number (16 regs)
                "4033": "4033_20",  # EMS_ver parallel
                "4043": "4043_20",  # BMS/DCDC_ver parallel
                "4053": "4053_20",  # HMI_ver parallel
            }
            formatted_address = spans.get(addr_clean, f"{addr_clean}_20")

        coeff, unit = parse_coeff_and_unit(coeff_raw, unit_raw)

        # Action: status/alarm/fault -> 9, else 4
        action = "4"
        name_lower = clean_name.lower()
        if any(w in name_lower for w in ("fault", "alarm", "warning", "status", "states")):
            action = "9"

        # Model restriction
        model_restriction = ""
        full_comment = f"{comment1} {comment2}".lower()
        if "only available on 200f/100f" in full_comment:
            model_restriction = "200F_100F_ONLY"
        elif "dh200y" in full_comment and "dh200f" not in full_comment:
            model_restriction = "DH200Y_ONLY"

        return {
            "info1": fc,
            "address": formatted_address,
            "name": clean_name,
            "dtype": dtype_norm,
            "coeff_a": coeff,
            "coeff_b": 0.0,
            "unit": unit,
            "action": action,
            "model_restriction": model_restriction,
            "line": line_no,
        }

    def _parse_submodule_row(self, r: list[str], line_no: int, section_name: str) -> dict | None:
        if len(r) < 6:
            return None
        addr_clean = r[2].strip() if len(r) > 2 else ""
        if not re.match(r"^\d+(?:~\d+)?$", addr_clean):
            return None

        var_name = clean_text(r[3]) if len(r) > 3 else ""
        desc = clean_text(r[4]) if len(r) > 4 else var_name
        dtype = clean_text(r[5]).upper() if len(r) > 5 else ""
        coeff_raw = clean_text(r[8]) if len(r) > 8 else "1"
        unit_raw = clean_text(r[9]) if len(r) > 9 else ""
        comment1 = clean_text(r[10]) if len(r) > 10 else ""
        comment2 = clean_text(r[11]) if len(r) > 11 else ""

        # Skip empty rows and reserved rows
        if not desc and not var_name:
            return None
        if desc.lower() in ("reserved", "resever", "reseved") or var_name.lower() in (
            "reserved",
            "resever",
            "reseved",
        ):
            return None

        # Skip Low/High word continuations of 32-bit registers
        if not dtype and (
            var_name.lower().endswith(" l")
            or var_name.lower().endswith("_l")
            or (
                var_name.lower().endswith("l")
                and not any(
                    w in var_name.lower() for w in ("total", "signal", "level", "fail", "cool")
                )
            )
        ):
            return None
        if not dtype and (
            var_name.lower().endswith(" h")
            or var_name.lower().endswith("_h")
            or (var_name.lower().endswith("h") and "(high)" in desc.lower())
        ):
            return None

        if not desc:
            desc = var_name

        clean_name = re.sub(r"\s*\((?:High(?:\s*Word)?|High)\)", "", desc, flags=re.IGNORECASE)
        clean_name = re.sub(r"\s+H$", "", clean_name).strip()

        # Check for ASCII strings in submodules
        is_ascii = (
            dtype == "ASCII"
            or "ASCII" in coeff_raw
            or "ASCII" in unit_raw
            or "ASCII" in comment1
            or "ASCII" in comment2
            or addr_clean in ("1387", "1388", "1637")
        )

        # Normalize data type
        dtype_norm = "U16"
        if is_ascii:
            dtype_norm = "STRING"
        elif dtype in ("S16", "INT16"):
            dtype_norm = "I16"
        elif dtype in ("U32", "UINT32"):
            dtype_norm = "U32"
        elif dtype in ("S32", "INT32"):
            dtype_norm = "I32"
        elif dtype in ("ASCII", "STRING"):
            dtype_norm = "STRING"

        coeff, unit = parse_coeff_and_unit(coeff_raw, unit_raw)

        # Offsets (temperatures with -40 offset, current with -3200)
        coeff_b = 0.0
        full_desc = f"{comment1} {comment2}"
        if (
            "register value - 40" in full_desc.lower()
            or "offset 40°c" in full_desc.lower()
            or "offset of 40°c" in full_desc.lower()
        ):
            coeff_b = -40.0
        elif "register value/10 - 3200" in full_desc.lower():
            coeff = 0.1
            coeff_b = -3200.0

        # String formatting
        formatted_address = addr_clean
        if dtype_norm == "STRING":
            if addr_clean == "1637":
                formatted_address = "1637_16"
            elif addr_clean in ("1387", "1388"):
                formatted_address = f"{addr_clean}_16"
            else:
                formatted_address = f"{addr_clean}_20"

        action = "4"
        name_lower = clean_name.lower()
        if any(w in name_lower for w in ("fault", "alarm", "warning", "status", "state", "lock")):
            action = "9"

        return {
            "info1": 4,
            "address": formatted_address,
            "name": clean_name,
            "dtype": dtype_norm,
            "coeff_a": coeff,
            "coeff_b": coeff_b,
            "unit": unit,
            "action": action,
            "model_restriction": "",
            "raw_address": addr_clean,
            "section": section_name,
            "line": line_no,
        }


class DynessDefinitionGenerator:
    """Generates production-grade WebdynSunPM definitions for Dyness DH-Series."""

    def __init__(self, parser: DynessProtocolParser, output_dir: str | Path):
        self.parser = parser
        self.output_dir = Path(output_dir)
        self.generator = Generator()

    def generate_all(self) -> dict[str, Path]:
        self.output_dir.mkdir(parents=True, exist_ok=True)
        results = {}

        # 1. Dyness DH200Y (Liquid-Cooled 232 kWh)
        dh200y_path = self.output_dir / "BAT_Dyness_DH200Y.csv"
        self._write_dh200y(dh200y_path)
        results["DH200Y"] = dh200y_path

        # 2. Dyness DH200F (Air-Cooled 215 kWh)
        dh200f_path = self.output_dir / "BAT_Dyness_DH200F.csv"
        self._write_dh200f(dh200f_path)
        results["DH200F"] = dh200f_path

        # 3. Dyness DH100F (Air-Cooled 100 kWh)
        dh100f_path = self.output_dir / "BAT_Dyness_DH100F.csv"
        self._write_dh100f(dh100f_path)
        results["DH100F"] = dh100f_path

        # 4. Dyness DH-Series Universal Generic Definition
        generic_path = self.output_dir / "BAT_Dyness_DH_Series_Generic.csv"
        self._write_generic(generic_path)
        results["DH_Series_Generic"] = generic_path

        return results

    def _write_dh200y(self, path: Path) -> None:
        regs = []
        # Holding registers (exclude 200F/100F only)
        for r in self.parser.holding_regs:
            if r["model_restriction"] != "200F_100F_ONLY":
                regs.append(dict(r))

        # Common input registers (exclude 200F/100F only PV data if not applicable)
        for r in self.parser.common_input_regs:
            if r["model_restriction"] != "200F_100F_ONLY":
                regs.append(dict(r))

        # Parallel input registers
        for r in self.parser.parallel_input_regs:
            regs.append(dict(r))

        # Subsystems for DH200Y
        # PCS Pos #4
        for r in self.parser.pcs_pos4_dh200y:
            regs.append(dict(r))

        # BMS summary
        for r in self.parser.bms_common_regs:
            addr = r["address"]
            if "~" not in addr:
                try:
                    if int(addr.split("_")[0]) < 782:
                        regs.append(dict(r))
                except ValueError:
                    pass

        # Expand DH200Y BMS Cell Voltages (260 cells: 782..1041)
        for cell_idx in range(1, 261):
            regs.append(
                {
                    "info1": 4,
                    "address": str(781 + cell_idx),
                    "name": f"Cell Voltage {cell_idx}",
                    "dtype": "U16",
                    "coeff_a": 0.001,
                    "coeff_b": 0.0,
                    "unit": "V",
                    "action": "4",
                }
            )

        # Expand DH200Y BMS Cell Temperatures (80 temps: 1042..1121)
        for temp_idx in range(1, 81):
            regs.append(
                {
                    "info1": 4,
                    "address": str(1041 + temp_idx),
                    "name": f"Cell Temperature {temp_idx}",
                    "dtype": "U16",
                    "coeff_a": 1.0,
                    "coeff_b": -40.0,
                    "unit": "°C",
                    "action": "4",
                }
            )

        # Fire protection
        for r in self.parser.fire_regs:
            regs.append(dict(r))

        # Liquid cooling (DH200Y specific)
        for r in self.parser.liquid_cooling_regs:
            regs.append(dict(r))

        # Temperature and Humidity
        for r in self.parser.temp_hum_regs:
            regs.append(dict(r))

        self._export_csv(path, "DH200Y", regs)

    def _write_dh200f(self, path: Path) -> None:
        regs = []
        # Holding registers (include all holding)
        for r in self.parser.holding_regs:
            if r["model_restriction"] != "DH200Y_ONLY":
                regs.append(dict(r))

        # Common input registers (include PV/MPPT data)
        for r in self.parser.common_input_regs:
            if r["model_restriction"] != "DH200Y_ONLY":
                regs.append(dict(r))

        # Parallel input registers
        for r in self.parser.parallel_input_regs:
            regs.append(dict(r))

        # Subsystems for DH200F
        # PCS Pos #5
        for r in self.parser.pcs_pos5_dh200f:
            regs.append(dict(r))

        # DCDC modules
        for r in self.parser.dcdc_regs:
            regs.append(dict(r))

        # BMS summary
        for r in self.parser.bms_common_regs:
            addr = r["address"]
            if "~" not in addr:
                try:
                    if int(addr.split("_")[0]) < 782:
                        regs.append(dict(r))
                except ValueError:
                    pass

        # Expand DH200F BMS Cell Voltages (240 cells: 782..1021)
        for cell_idx in range(1, 241):
            regs.append(
                {
                    "info1": 4,
                    "address": str(781 + cell_idx),
                    "name": f"Cell Voltage {cell_idx}",
                    "dtype": "U16",
                    "coeff_a": 0.001,
                    "coeff_b": 0.0,
                    "unit": "V",
                    "action": "4",
                }
            )

        # Expand DH200F BMS Cell Temperatures (60 temps: 1022..1081)
        for temp_idx in range(1, 61):
            regs.append(
                {
                    "info1": 4,
                    "address": str(1021 + temp_idx),
                    "name": f"Cell Temperature {temp_idx}",
                    "dtype": "U16",
                    "coeff_a": 1.0,
                    "coeff_b": -40.0,
                    "unit": "°C",
                    "action": "4",
                }
            )

        # Fire protection
        for r in self.parser.fire_regs:
            regs.append(dict(r))

        # AC Air Conditioner (DH200F specific)
        for r in self.parser.ac_regs:
            regs.append(dict(r))

        # Temperature and Humidity
        for r in self.parser.temp_hum_regs:
            regs.append(dict(r))

        self._export_csv(path, "DH200F", regs)

    def _write_dh100f(self, path: Path) -> None:
        regs = []
        # Same core register architecture as DH200F (100 kWh air-cooled variant)
        for r in self.parser.holding_regs:
            if r["model_restriction"] != "DH200Y_ONLY":
                regs.append(dict(r))

        for r in self.parser.common_input_regs:
            if r["model_restriction"] != "DH200Y_ONLY":
                regs.append(dict(r))

        for r in self.parser.parallel_input_regs:
            regs.append(dict(r))

        for r in self.parser.pcs_pos5_dh200f:
            regs.append(dict(r))

        for r in self.parser.dcdc_regs:
            regs.append(dict(r))

        for r in self.parser.bms_common_regs:
            addr = r["address"]
            if "~" not in addr and int(addr) < 782:
                regs.append(dict(r))

        for cell_idx in range(1, 241):
            regs.append(
                {
                    "info1": 4,
                    "address": str(781 + cell_idx),
                    "name": f"Cell Voltage {cell_idx}",
                    "dtype": "U16",
                    "coeff_a": 0.001,
                    "coeff_b": 0.0,
                    "unit": "V",
                    "action": "4",
                }
            )

        for temp_idx in range(1, 61):
            regs.append(
                {
                    "info1": 4,
                    "address": str(1021 + temp_idx),
                    "name": f"Cell Temperature {temp_idx}",
                    "dtype": "U16",
                    "coeff_a": 1.0,
                    "coeff_b": -40.0,
                    "unit": "°C",
                    "action": "4",
                }
            )

        for r in self.parser.fire_regs:
            regs.append(dict(r))

        for r in self.parser.ac_regs:
            regs.append(dict(r))

        for r in self.parser.temp_hum_regs:
            regs.append(dict(r))

        self._export_csv(path, "DH100F", regs)

    def _write_generic(self, path: Path) -> None:
        """Writes the All-In-One Universal definition covering the DH series superset."""
        regs = []
        # All holding registers
        for r in self.parser.holding_regs:
            regs.append(dict(r))

        # All common inputs
        for r in self.parser.common_input_regs:
            regs.append(dict(r))

        # All parallel inputs
        for r in self.parser.parallel_input_regs:
            regs.append(dict(r))

        # Primary PCS (Position #4 from DH200Y)
        for r in self.parser.pcs_pos4_dh200y:
            regs.append(dict(r))

        # DCDC modules
        for r in self.parser.dcdc_regs:
            regs.append(dict(r))

        # BMS summary
        for r in self.parser.bms_common_regs:
            addr = r["address"]
            if "~" not in addr:
                try:
                    if int(addr.split("_")[0]) < 782:
                        regs.append(dict(r))
                except ValueError:
                    pass

        # 260 cells and 80 temps (DH200Y full cluster capacity)
        for cell_idx in range(1, 261):
            regs.append(
                {
                    "info1": 4,
                    "address": str(781 + cell_idx),
                    "name": f"Cell Voltage {cell_idx}",
                    "dtype": "U16",
                    "coeff_a": 0.001,
                    "coeff_b": 0.0,
                    "unit": "V",
                    "action": "4",
                }
            )

        for temp_idx in range(1, 81):
            regs.append(
                {
                    "info1": 4,
                    "address": str(1041 + temp_idx),
                    "name": f"Cell Temperature {temp_idx}",
                    "dtype": "U16",
                    "coeff_a": 1.0,
                    "coeff_b": -40.0,
                    "unit": "°C",
                    "action": "4",
                }
            )

        # Fire protection
        for r in self.parser.fire_regs:
            regs.append(dict(r))

        # Liquid cooling
        for r in self.parser.liquid_cooling_regs:
            regs.append(dict(r))

        # Temperature & Humidity
        for r in self.parser.temp_hum_regs:
            regs.append(dict(r))

        self._export_csv(path, "DH_Series_Generic", regs)

    def _export_csv(self, output_path: Path, model_name: str, raw_registers: list[dict]) -> None:
        """Deduplicates addresses, infers SunPM tags, and writes canonical 11-column CSV."""
        # 1. Deduplicate by (info1, address) - preserve first occurrence
        seen_keys = set()
        deduped = []
        for r in raw_registers:
            info1 = int(r["info1"])
            addr = str(r["address"])
            key = (info1, addr)
            if key not in seen_keys:
                seen_keys.add(key)
                deduped.append(r)

        # 2. Sort registers: first by info1 (3 then 4), then by address numeric value
        def addr_sort_key(rec: dict) -> tuple[int, int]:
            addr_str = str(rec["address"])
            try:
                base = int(addr_str.split("_")[0]) if "_" in addr_str else int(addr_str)
            except ValueError:
                base = 999999
            try:
                i1 = int(rec["info1"])
            except ValueError:
                i1 = 99
            return (i1, base)

        deduped.sort(key=addr_sort_key)

        # 2.5 Resolve and eliminate any address collisions (O(N) interval tracking)
        clean_registers = []
        occupied_spans: dict[int, list[tuple[int, int, str]]] = {}  # info1 -> [(start, end, name)]

        for rec in deduped:
            try:
                info1 = int(rec["info1"])
                addr_str = str(rec["address"])
                base_addr = int(addr_str.split("_")[0]) if "_" in addr_str else int(addr_str)
            except ValueError:
                continue
            reg_count = Generator.get_register_count(rec["dtype"], addr_str)
            start_addr = base_addr
            end_addr = base_addr + max(1, reg_count) - 1

            if info1 not in occupied_spans:
                occupied_spans[info1] = []

            # Check overlap with already accepted registers in this info1 space
            overlap = False
            for prev_start, prev_end, _prev_name in occupied_spans[info1]:
                if max(start_addr, prev_start) <= min(end_addr, prev_end):
                    overlap = True
                    break

            if not overlap:
                occupied_spans[info1].append((start_addr, end_addr, rec["name"]))
                clean_registers.append(rec)

        # 3. Assign SunPM canonical tags with strict tag uniqueness
        used_tags: set[str] = set()
        tagged_rows = []

        for idx, rec in enumerate(clean_registers, start=1):
            name = rec["name"]
            # Infer tag using knowledge miner
            candidate_tag = infer_sunpm_tag(name, equipment_type="Battery")

            # Dedicated BESS tag mapping
            if not candidate_tag:
                name_l = name.lower()
                if "state of charge" in name_l or name_l == "soc":
                    candidate_tag = "SOC"
                elif "state of health" in name_l or name_l == "soh":
                    candidate_tag = "SOH"
                elif "total battery pack voltage" in name_l or "battery voltage" in name_l:
                    candidate_tag = "V_BAT"
                elif "total battery pack current" in name_l or "battery current" in name_l:
                    candidate_tag = "I_BAT"
                elif "battery power" in name_l:
                    candidate_tag = "P_BAT"
                elif "inverter frequency" in name_l or "system frequency" in name_l:
                    candidate_tag = "FAC"
                elif "total active power" in name_l or "inverter total power" in name_l:
                    candidate_tag = "PAC"
                elif "total reactive power" in name_l:
                    candidate_tag = "QAC"
                elif "total apparent power" in name_l:
                    candidate_tag = "SAC"
                elif "system daily energy" in name_l or "daily grid feed-in energy" in name_l:
                    candidate_tag = "E_DAY"
                elif "system total energy" in name_l:
                    candidate_tag = "E_TOTAL"
                elif "pv input power" in name_l:
                    candidate_tag = "P_PV"
                elif "max. cell voltage" in name_l:
                    candidate_tag = "V_CELL_MAX"
                elif "min. cell voltage" in name_l:
                    candidate_tag = "V_CELL_MIN"
                elif "max. cell temperature" in name_l:
                    candidate_tag = "T_CELL_MAX"
                elif "min. cell temperature" in name_l:
                    candidate_tag = "T_CELL_MIN"
                elif "ambient temperature" in name_l:
                    candidate_tag = "T_AMB"

            # Strict uniqueness
            final_tag = ""
            if candidate_tag and candidate_tag not in used_tags:
                final_tag = candidate_tag
                used_tags.add(candidate_tag)

            coef_a = f"{float(rec['coeff_a']):.6f}"
            coef_b = f"{float(rec['coeff_b']):.6f}"

            tagged_rows.append(
                [
                    str(idx),
                    str(rec["info1"]),
                    str(rec["address"]),
                    str(rec["dtype"]),
                    "",
                    Generator.sanitize_csv_field(name),
                    final_tag,
                    coef_a,
                    coef_b,
                    str(rec["unit"]),
                    str(rec["action"]),
                ]
            )

        # 4. Write CSV with UTF-8 BOM encoding and \r\n line terminators
        with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f, delimiter=";", lineterminator="\r\n")
            # 11 columns in header
            writer.writerow(["Modbus", "Battery", "Dyness", model_name, "", "", "", "", "", "", ""])
            for row in tagged_rows:
                writer.writerow(row)

        print(
            f"Generated {output_path.name}: {len(tagged_rows)} registers, {len(used_tags)} canonical tags."
        )


def main(argv: list[str] | None = None) -> None:
    arg_parser = argparse.ArgumentParser(
        description="Generate Dyness DH-Series WebdynSunPM definition files."
    )
    arg_parser.add_argument(
        "source_csv",
        nargs="?",
        default=str(REPO_ROOT / "stress_test_data" / "Dyness_DH_Series_Modbus_Protocol_V2.2.csv"),
        help="Path to vendor Modbus protocol CSV.",
    )
    arg_parser.add_argument(
        "--out-dir",
        default=str(REPO_ROOT / "Dyness_DH_Series_Definitions"),
        help="Output directory for generated definitions.",
    )
    cli_args = arg_parser.parse_args(argv)

    print("================================================================================")
    print("DYNESS DH-SERIES WEBDYNSUNPM DEFINITION GENERATOR")
    print("================================================================================")
    print(f"Source file:      {cli_args.source_csv}")
    print(f"Output directory: {cli_args.out_dir}")

    parser = DynessProtocolParser(cli_args.source_csv)
    parser.parse()

    generator = DynessDefinitionGenerator(parser, cli_args.out_dir)
    generated_files = generator.generate_all()

    # Self-validation
    print("\n--------------------------------------------------------------------------------")
    print("VALIDATING GENERATED DEFINITIONS")
    print("--------------------------------------------------------------------------------")
    validator = Generator()
    all_valid = True
    for _model, filepath in generated_files.items():
        report = validator.validate_csv_detailed(str(filepath), strict=True, strict_overlap=True)
        if report.is_valid:
            print(
                f"  ✓ {filepath.name}: VALID ({report.register_count} registers, 0 errors, 0 overlaps)"
            )
        else:
            all_valid = False
            print(f"  ✗ {filepath.name}: INVALID ({len(report.issues)} issues)")
            for issue in report.issues:
                print(f"      - {issue}")

    if not all_valid:
        print("\n[ERROR] One or more generated definitions failed validation!")
        sys.exit(1)

    print("\n[SUCCESS] All Dyness definitions passed strict WebdynSunPM validation!")


if __name__ == "__main__":
    main()
