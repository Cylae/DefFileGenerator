#!/usr/bin/env python3
"""
Dyness DH-Series Commercial & Industrial BESS Definition Generator.
(DH200Y, DH200F, DH100F, and DH_Series_Generic).

Strictly adheres to WebdynSunPM reference specifications:
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
from typing import Any

# Ensure parent directory and .venv are in sys.path to support direct and debug executions
REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

# Auto-locate virtual environment site-packages if run with system python in VS Code
for candidate_dir in [REPO_ROOT, Path(__file__).resolve().parent]:
    win_sp = candidate_dir / ".venv" / "Lib" / "site-packages"
    if win_sp.is_dir() and str(win_sp) not in sys.path:
        sys.path.insert(0, str(win_sp))
        break

from DefFileGenerator.def_gen import Generator  # noqa: E402
from DefFileGenerator.knowledge_miner import infer_sunpm_tag  # noqa: E402

__all__ = [
    "DynessDefinitionGenerator",
    "DynessProtocolParser",
    "clean_text",
    "generate_dyness_definitions",
    "parse_coeff_and_unit",
]


def clean_text(val: Any) -> str:
    """Strips whitespace and neutralizes quotes."""
    if val is None:
        return ""
    s = str(val).strip()
    return s.replace('"', "")


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
    if final_unit in ("℃", "C", "° C", "0.1℃", "degC", "celsius"):
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
        self.holding_regs: list[dict[str, Any]] = []
        self.common_input_regs: list[dict[str, Any]] = []
        self.parallel_input_regs: list[dict[str, Any]] = []
        self.pcs_pos4_dh200y: list[dict[str, Any]] = []
        self.pcs_pos5_dh200f: list[dict[str, Any]] = []
        self.dcdc_regs: list[dict[str, Any]] = []
        self.bms_common_regs: list[dict[str, Any]] = []
        self.fire_regs: list[dict[str, Any]] = []
        self.ac_regs: list[dict[str, Any]] = []
        self.liquid_cooling_regs: list[dict[str, Any]] = []
        self.temp_hum_regs: list[dict[str, Any]] = []

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

    def _parse_holding_row(self, r: list[str], line_no: int) -> dict[str, Any] | None:
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

    def _parse_standard_input_row(
        self, r: list[str], line_no: int, fc: int
    ) -> dict[str, Any] | None:
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

        if not dtype and not is_ascii:
            return None

        if not desc:
            desc = var_name

        clean_name = re.sub(r"\s*\((?:High(?:\s*Word)?|High)\)", "", desc, flags=re.IGNORECASE)
        clean_name = re.sub(r"\s+H$", "", clean_name).strip()

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

        formatted_address = addr_clean
        if dtype_norm == "STRING":
            spans = {
                "33": "33_20",
                "43": "43_20",
                "53": "53_20",
                "254": "254_24",
                "2125": "2125_20",
                "2135": "2135_20",
                "2145": "2145_30",
                "2160": "2160_20",
                "2170": "2170_32",
                "4033": "4033_20",
                "4043": "4043_20",
                "4053": "4053_20",
            }
            formatted_address = spans.get(addr_clean, f"{addr_clean}_20")

        coeff, unit = parse_coeff_and_unit(coeff_raw, unit_raw)

        action = "4"
        name_lower = clean_name.lower()
        if any(w in name_lower for w in ("fault", "alarm", "warning", "status", "states")):
            action = "9"

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

    def _parse_submodule_row(
        self, r: list[str], line_no: int, section_name: str
    ) -> dict[str, Any] | None:
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

        if not desc and not var_name:
            return None
        if desc.lower() in ("reserved", "resever", "reseved") or var_name.lower() in (
            "reserved",
            "resever",
            "reseved",
        ):
            return None

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

        is_ascii = (
            dtype == "ASCII"
            or "ASCII" in coeff_raw
            or "ASCII" in unit_raw
            or "ASCII" in comment1
            or "ASCII" in comment2
            or addr_clean in ("1387", "1388", "1637")
        )

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
        results: dict[str, Path] = {}

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

    @staticmethod
    def _filter_bms_summary(
        bms_regs: list[dict[str, Any]], max_addr: int = 782
    ) -> list[dict[str, Any]]:
        """Filters summary BMS registers below the expanded cell registers threshold."""
        filtered = []
        for r in bms_regs:
            addr = str(r["address"])
            if "~" not in addr:
                try:
                    base = int(addr.split("_")[0])
                    if base < max_addr:
                        filtered.append(dict(r))
                except ValueError:
                    pass
        return filtered

    @staticmethod
    def _expand_cells(
        start_addr: int, count: int, name_prefix: str, coeff_a: float, coeff_b: float, unit: str
    ) -> list[dict[str, Any]]:
        """Expands individual battery cell voltage or temperature registers."""
        expanded: list[dict[str, Any]] = []
        for idx in range(1, count + 1):
            expanded.append(
                {
                    "info1": 4,
                    "address": str(start_addr + idx),
                    "name": f"{name_prefix} {idx}",
                    "dtype": "U16",
                    "coeff_a": coeff_a,
                    "coeff_b": coeff_b,
                    "unit": unit,
                    "action": "4",
                }
            )
        return expanded

    def _write_dh200y(self, path: Path) -> None:
        regs: list[dict[str, Any]] = []
        # Holding registers (exclude 200F/100F only)
        for r in self.parser.holding_regs:
            if r["model_restriction"] != "200F_100F_ONLY":
                regs.append(dict(r))

        # Common input registers
        for r in self.parser.common_input_regs:
            if r["model_restriction"] != "200F_100F_ONLY":
                regs.append(dict(r))

        # Parallel input registers
        for r in self.parser.parallel_input_regs:
            regs.append(dict(r))

        # Subsystems for DH200Y
        for r in self.parser.pcs_pos4_dh200y:
            regs.append(dict(r))

        # BMS summary
        regs.extend(self._filter_bms_summary(self.parser.bms_common_regs, max_addr=782))

        # Expand DH200Y BMS Cell Voltages (260 cells: 782..1041)
        regs.extend(
            self._expand_cells(
                start_addr=781,
                count=260,
                name_prefix="Cell Voltage",
                coeff_a=0.001,
                coeff_b=0.0,
                unit="V",
            )
        )

        # Expand DH200Y BMS Cell Temperatures (80 temps: 1042..1121)
        regs.extend(
            self._expand_cells(
                start_addr=1041,
                count=80,
                name_prefix="Cell Temperature",
                coeff_a=1.0,
                coeff_b=-40.0,
                unit="°C",
            )
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
        regs: list[dict[str, Any]] = []
        # Holding registers
        for r in self.parser.holding_regs:
            if r["model_restriction"] != "DH200Y_ONLY":
                regs.append(dict(r))

        # Common input registers
        for r in self.parser.common_input_regs:
            if r["model_restriction"] != "DH200Y_ONLY":
                regs.append(dict(r))

        # Parallel input registers
        for r in self.parser.parallel_input_regs:
            regs.append(dict(r))

        # Subsystems for DH200F
        for r in self.parser.pcs_pos5_dh200f:
            regs.append(dict(r))

        for r in self.parser.dcdc_regs:
            regs.append(dict(r))

        # BMS summary
        regs.extend(self._filter_bms_summary(self.parser.bms_common_regs, max_addr=782))

        # Expand DH200F BMS Cell Voltages (240 cells: 782..1021)
        regs.extend(
            self._expand_cells(
                start_addr=781,
                count=240,
                name_prefix="Cell Voltage",
                coeff_a=0.001,
                coeff_b=0.0,
                unit="V",
            )
        )

        # Expand DH200F BMS Cell Temperatures (60 temps: 1022..1081)
        regs.extend(
            self._expand_cells(
                start_addr=1021,
                count=60,
                name_prefix="Cell Temperature",
                coeff_a=1.0,
                coeff_b=-40.0,
                unit="°C",
            )
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
        regs: list[dict[str, Any]] = []
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

        # BMS summary
        regs.extend(self._filter_bms_summary(self.parser.bms_common_regs, max_addr=782))

        # Expand DH100F BMS Cell Voltages (240 cells: 782..1021)
        regs.extend(
            self._expand_cells(
                start_addr=781,
                count=240,
                name_prefix="Cell Voltage",
                coeff_a=0.001,
                coeff_b=0.0,
                unit="V",
            )
        )

        # Expand DH100F BMS Cell Temperatures (60 temps: 1022..1081)
        regs.extend(
            self._expand_cells(
                start_addr=1021,
                count=60,
                name_prefix="Cell Temperature",
                coeff_a=1.0,
                coeff_b=-40.0,
                unit="°C",
            )
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
        regs: list[dict[str, Any]] = []
        for r in self.parser.holding_regs:
            regs.append(dict(r))

        for r in self.parser.common_input_regs:
            regs.append(dict(r))

        for r in self.parser.parallel_input_regs:
            regs.append(dict(r))

        for r in self.parser.pcs_pos4_dh200y:
            regs.append(dict(r))

        for r in self.parser.dcdc_regs:
            regs.append(dict(r))

        # BMS summary
        regs.extend(self._filter_bms_summary(self.parser.bms_common_regs, max_addr=782))

        # 260 cells and 80 temps (DH200Y full cluster capacity)
        regs.extend(
            self._expand_cells(
                start_addr=781,
                count=260,
                name_prefix="Cell Voltage",
                coeff_a=0.001,
                coeff_b=0.0,
                unit="V",
            )
        )

        regs.extend(
            self._expand_cells(
                start_addr=1041,
                count=80,
                name_prefix="Cell Temperature",
                coeff_a=1.0,
                coeff_b=-40.0,
                unit="°C",
            )
        )

        for r in self.parser.fire_regs:
            regs.append(dict(r))

        for r in self.parser.liquid_cooling_regs:
            regs.append(dict(r))

        for r in self.parser.temp_hum_regs:
            regs.append(dict(r))

        self._export_csv(path, "DH_Series_Generic", regs)

    def _export_csv(
        self, output_path: Path, model_name: str, raw_registers: list[dict[str, Any]]
    ) -> None:
        """Deduplicates addresses, infers SunPM tags, and writes canonical 11-column CSV."""
        # 1. Deduplicate by (info1, address) - preserve first occurrence
        seen_keys: set[tuple[int, str]] = set()
        deduped: list[dict[str, Any]] = []
        for r in raw_registers:
            info1 = int(r["info1"])
            addr = str(r["address"])
            key = (info1, addr)
            if key not in seen_keys:
                seen_keys.add(key)
                deduped.append(r)

        # 2. Sort registers: first by info1 (3 then 4), then by address numeric value
        def addr_sort_key(rec: dict[str, Any]) -> tuple[int, int]:
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
        clean_registers: list[dict[str, Any]] = []
        occupied_spans: dict[int, list[tuple[int, int, str]]] = {}

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
        tagged_rows: list[list[str]] = []

        for idx, rec in enumerate(clean_registers, start=1):
            name = str(rec["name"])
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
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f, delimiter=";", lineterminator="\r\n")
            # 11 columns in header
            writer.writerow(["Modbus", "Battery", "Dyness", model_name, "", "", "", "", "", "", ""])
            for row in tagged_rows:
                writer.writerow(row)


def generate_dyness_definitions(
    source_csv: str | Path,
    out_dir: str | Path,
) -> dict[str, Path]:
    """
    Parses vendor Modbus protocol CSV and generates strictly valid WebdynSunPM definitions.

    Returns:
        dict[str, Path]: Mapping of model name to generated CSV Path.
    """
    parser = DynessProtocolParser(source_csv)
    parser.parse()
    generator = DynessDefinitionGenerator(parser, out_dir)
    return generator.generate_all()


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

    print("=" * 80)
    print("DYNESS DH-SERIES WEBDYNSUNPM DEFINITION GENERATOR")
    print("=" * 80)
    print(f"Source file:      {cli_args.source_csv}")
    print(f"Output directory: {cli_args.out_dir}")

    generated_files = generate_dyness_definitions(cli_args.source_csv, cli_args.out_dir)

    print("\n" + "-" * 80)
    print("VALIDATING GENERATED DEFINITIONS")
    print("-" * 80)
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

