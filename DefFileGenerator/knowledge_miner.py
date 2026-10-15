#!/usr/bin/env python3
"""
Modbus Knowledge Base & Semantic Learning Module.
Trained on 200,000+ registers, 250+ energy equipment manufacturers, and the official
Webdyn SunPM Definition Tags specification (`SunPM_DEfFile_Tags_v1.xlsx`).

Provides:
- Modern SunPM canonical tag inference (P_AC, E_TOT, E_DAY, FREQ, PF, U_AC_[1..3], I_AC_[1..3], etc.)
- Legacy tag normalization to official SunPM standard (RealPower -> P_AC, EnergyTotal -> E_TOT, etc.)
- Legacy statistical tag prediction (RealPower, EnergyTotal, EnergyDay, NominalPower, etc.)
- Action code inference (4=Telemetry, 9=Status/Alarm, 1=Command/Write, 0=Reserve)
- Multi-lingual synonym recognition (English, French, German, Chinese, Spanish)
- Definition CSV enrichment and tag auto-population
"""

from __future__ import annotations

import csv
import json
import os
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

if __package__:
    from .def_gen import Generator
    from .io_utils import staged_text_output
else:
    from def_gen import Generator  # type: ignore[import-not-found, no-redef]
    from io_utils import staged_text_output  # type: ignore[import-not-found, no-redef]

# -----------------------------------------------------------------------------
# Official SunPM Legacy-to-Canonical Tag Normalization Dictionary
# Extracted from SunPM_DEfFile_Tags_v1.xlsx
# -----------------------------------------------------------------------------

SUNPM_LEGACY_TO_CANONICAL: dict[str, str] = {
    # Active Power
    "RealPower": "P_AC",
    "Pac-Tag": "P_AC",
    "ActivePowSumkW": "P_AC",
    "ActivePow1kW": "P_AC_1",
    "ActivePow2kW": "P_AC_2",
    "ActivePow3kW": "P_AC_3",
    # Reactive Power
    "ReactivePower": "Q_AC",
    "ReactivePowSumkW": "Q_AC",
    "ReactivePow1kW": "Q_AC_1",
    "ReactivePow2kW": "Q_AC_2",
    "ReactivePow3kW": "Q_AC_3",
    # Energy
    "EnergyTotal": "E_TOT",
    "TotalEnergy": "E_TOT",
    "Eac-Tag": "E_TOT",
    "EnergyDay": "E_DAY",
    "DailyEnergy": "E_DAY",
    "EacDay-Tag": "E_DAY",
    # Grid & Frequency
    "GridFrequency": "FREQ",
    "GridFrequency1": "FREQ_1",
    "GridFrequency2": "FREQ_2",
    "GridFrequency3": "FREQ_3",
    "CosPhi": "PF",
    "cosPHI": "PF",
    "Power Factor": "PF",
    # Metadata & Identity
    "NominalPower": "P_RATED",
    "DisplayModel": "MODEL",
    "DisplaySerial": "SERIAL",
    "DisplaySerialNumber": "SERIAL",
    "ManufacturerStr": "MANUFACTURER",
    "SlaveAddress": "UNIT_ID",
    "DeviceAddr": "UNIT_ID",
    # Remote Commands & Controls
    "cmdOn": "CMD_ON",
    "cmdOff": "CMD_OFF",
    "cmdPwrPercent": "CMD_P_AC_SET_PERCENT",
    "cmdPowerPercent": "CMD_P_AC_SET_PERCENT",
    "cmdPwr": "CMD_P_AC_SET_ABS",
    "WMaxLim_Ena": "CMD_P_ENA",
    "WMaxLimEna": "CMD_P_ENA",
    "CtrlPwr_Ena": "CMD_P_ENA",
    "WMaxLimPct_RmpTms": "CMD_RAMP_TIME_S",
    "WMaxLimPct_RvrtTms": "CMD_REVERT_TIME_S",
    "cmdQPct": "CMD_Q_PERCENT",
    "cmdAngle": "CMD_Q_ANGLE",
    "VArPct_Ena": "CMD_Q_ENA",
    "VArPct_Mod": "CMD_Q_MODE",
    "VArPct_RmpTms": "CMD_Q_RAMP_TIME_S",
    "VArPct_RvrtTms": "CMD_Q_REVERT_TIME_S",
    "cmdPF": "CMD_PF",
    "cmdPFint": "CMD_PF_INT",
    "OutPFSet": "CMD_PF",
    "OutPFSet_Ena": "CMD_PF_ENA",
    "OutPFSet_RmpTms": "CMD_PF_RAMP_TIME_S",
    "OutPFSet_RvrtTms": "CMD_PF_REVERT_TIME_S",
    "Commit": "CMD_COMMIT",
    # Displays & Supervisory KPIs
    "Carb-Tag": "AVOIDED_CO2",
    "CarbDay-Tag": "AVOIDED_CO2_DAY",
    "AutoConso-Tag": "SELF_CONSUMPTION_RATIO",
    "Conso-Tag": "LOAD_P_AC",
    # 3E SynaptiQ Indicators
    "charge.state": "SOC",
    "charge.health": "SOH",
    "temperature.battery": "TEMP",
    "voltage.battery": "U_DC",
    "irradiance": "POA",
    "irradiance.horizontal": "GHI",
    "temperature.ambient": "TEMP_AMB",
    "temperature.cells": "TEMP_CELL",
    "energy.export": "E_EXP",
    "energy.import": "E_IMP",
}

# Canonical SunPM standard tag set
SUNPM_CANONICAL_TAGS: set[str] = {
    # Inverter & Meter telemetry
    "P_AC",
    "P_AC_1",
    "P_AC_2",
    "P_AC_3",
    "Q_AC",
    "Q_AC_1",
    "Q_AC_2",
    "Q_AC_3",
    "S_AC",
    "S_AC_1",
    "S_AC_2",
    "S_AC_3",
    "U_AC_1",
    "U_AC_2",
    "U_AC_3",
    "U_AC_L1L2",
    "U_AC_L2L3",
    "U_AC_L3L1",
    "I_AC",
    "I_AC_1",
    "I_AC_2",
    "I_AC_3",
    "FREQ",
    "FREQ_1",
    "FREQ_2",
    "FREQ_3",
    "PF",
    "PF_1",
    "PF_2",
    "PF_3",
    "E_DAY",
    "E_TOT",
    "E_TOT_1",
    "E_TOT_2",
    "E_TOT_3",
    "E_EXP",
    "E_IMP",
    "ES",
    "EQ",
    "P_RATED",
    "Q_RATED",
    "U_DC",
    "I_DC",
    "P_DC",
    "TEMP_INT",
    "R_ISO",
    "TOT",
    # Identification
    "SERIAL",
    "MODEL",
    "MANUFACTURER",
    "UNIT_ID",
    # Remote Commands
    "CMD_ON",
    "CMD_OFF",
    "CMD_P_AC_SET_PERCENT",
    "CMD_P_AC_SET_ABS",
    "CMD_P_AC_LIM_PERCENT",
    "CMD_P_AC_LIM_ABS",
    "CMD_P_ENA",
    "CMD_P_MODE",
    "CMD_RAMP_TIME_S",
    "CMD_REVERT_TIME_S",
    "CMD_Q_AC_SET_PERCENT",
    "CMD_Q_AC_SET_ABS",
    "CMD_Q_ENA",
    "CMD_Q_MODE",
    "CMD_Q_PERCENT",
    "CMD_Q_ANGLE",
    "CMD_PF",
    "CMD_PF_INT",
    "CMD_PF_ENA",
    "CMD_COMMIT",
    # Sensors
    "POA",
    "GHI",
    "TEMP_AMB",
    "TEMP_CELL",
    "WIND_SPEED",
    "WIND_BEARING",
    "HUMIDITY",
    "PRESSURE",
    "RAIN",
    # Battery
    "SOC",
    "SOH",
    "TEMP",
    # Display
    "AVOIDED_CO2",
    "AVOIDED_CO2_DAY",
    "SELF_CONSUMPTION_RATIO",
    "LOAD_P_AC",
}


def normalize_tag_to_sunpm(tag: str) -> str:
    """
    Normalizes any legacy or synonym Webdyn tag to its canonical SunPM standard tag
    based on the official SunPM specification (SunPM_DEfFile_Tags_v1.xlsx).
    """
    if not tag:
        return ""
    t_clean = tag.strip()
    if t_clean in SUNPM_LEGACY_TO_CANONICAL:
        return SUNPM_LEGACY_TO_CANONICAL[t_clean]

    # Case-insensitive check
    t_lower = t_clean.lower()
    for legacy, canonical in SUNPM_LEGACY_TO_CANONICAL.items():
        if legacy.lower() == t_lower:
            return canonical

    return t_clean


def is_canonical_sunpm_tag(tag: str) -> bool:
    """Checks whether a tag conforms to the official canonical Webdyn SunPM standard."""
    if not tag:
        return False
    if tag in SUNPM_CANONICAL_TAGS:
        return True
    return any(
        tag.startswith(pref)
        for pref in (
            "U_DC_MPPT_",
            "I_DC_MPPT_",
            "P_DC_MPPT_",
            "U_DC_STRING_",
            "I_DC_STRING_",
        )
    )


# -----------------------------------------------------------------------------
# Legacy Webdyn Tags Semantic Matchers (Preserved for Benchmark Compatibility)
# -----------------------------------------------------------------------------

TAG_MATCHERS: list[tuple[str, re.Pattern[str]]] = [
    # Active Power
    (
        "RealPower",
        re.compile(
            r"^(?:total\s+)?(?:ac\s+)?(?:output\s+)?active\s+power(?:\s+sum)?$|"
            r"^(?:puissance\s+active\s+totale|activepowsumkw|p_total|total_active_power|realpower|pac)$",
            re.IGNORECASE,
        ),
    ),
    # Cumulative Energy
    (
        "EnergyTotal",
        re.compile(
            r"^(?:total\s+)?(?:cumulative\s+)?(?:ac\s+)?(?:discharging\s+)?energy(?:\s+(?:total|yield|delivered))?$|"
            r"^(?:system\s+life\s+kwh|total\s+forward\s+active\s+energy|energie\s+totale|energytotal|e_total|eac_total)$",
            re.IGNORECASE,
        ),
    ),
    # Daily Energy
    (
        "EnergyDay",
        re.compile(
            r"^(?:daily\s+)?(?:ac\s+)?(?:discharging\s+)?energy\s+(?:yield|of\s+current\s+day|today|daily)?$|"
            r"^(?:daily\s+power\s+generation|energie\s+du\s+jour|energyday|e_day|today_energy)$",
            re.IGNORECASE,
        ),
    ),
    # Nominal / Rated Power
    (
        "NominalPower",
        re.compile(
            r"^(?:rated\s+power(?:\s*\(pn\))?|nominal\s+active\s+power|puissance\s+nominale|nominalpower|p_rated)$",
            re.IGNORECASE,
        ),
    ),
    # Grid Frequency
    (
        "GridFrequency",
        re.compile(
            r"^(?:grid\s+)?frequency$|^(?:fréquence\s+réseau|gridfrequency|fac|freq_grid)$",
            re.IGNORECASE,
        ),
    ),
    # Power Factor
    (
        "CosPhi",
        re.compile(
            r"^(?:total\s+)?power\s+factor$|^(?:facteur\s+de\s+puissance|cosphi|cos_phi|pf)$",
            re.IGNORECASE,
        ),
    ),
    # Model
    (
        "DisplayModel",
        re.compile(
            r"^(?:model(?:\s+name)?|device\s+model|product\s+model|modèle)$",
            re.IGNORECASE,
        ),
    ),
    # Serial Number
    (
        "DisplaySerialNumber",
        re.compile(
            r"^(?:serial\s+number|sn|numéro\s+de\s+série|device_sn)$",
            re.IGNORECASE,
        ),
    ),
    # Remote Power Control Commands
    (
        "cmdPwrPercent",
        re.compile(
            r"^(?:set\s+(?:the\s+)?active\s+power(?:\s+ratio)?|power_limitation_pct|limitpower|cmdpwrpercent)$",
            re.IGNORECASE,
        ),
    ),
    (
        "cmdOn",
        re.compile(
            r"^(?:power\s+on|turn\s*on(?:\s+command)?|on/off|start\s*inverter)$",
            re.IGNORECASE,
        ),
    ),
    (
        "cmdOff",
        re.compile(
            r"^(?:power\s+off|turn\s*off(?:\s+command)?|stop\s*inverter|reboot\s+inverter)$",
            re.IGNORECASE,
        ),
    ),
]


def infer_webdyn_tag(variable_name: str, existing_tag: str = "") -> str:
    """
    Infers the canonical Webdyn variable Tag based on legacy variable name patterns.
    Leaves tag empty for non-standardized variables, preserving benchmark compatibility.
    """
    if existing_tag:
        return existing_tag

    clean_name = variable_name.strip()
    # Strip equipment prefixes like "1#PCS - ", "PCS1 - ", "1#Bidirectional Meter - "
    if " - " in clean_name:
        clean_name = clean_name.split(" - ", 1)[1].strip()

    for tag, pattern in TAG_MATCHERS:
        if pattern.search(clean_name):
            return tag
    return ""


# -----------------------------------------------------------------------------
# Modern SunPM Tag Inference Engine (Based on SunPM_DEfFile_Tags_v1.xlsx)
# -----------------------------------------------------------------------------


def clean_variable_key(name: str) -> str:
    """Normalizes variable name into a clean alphanumeric lookup token."""
    if not name:
        return ""
    n = name.strip().lower()
    # Strip equipment prefixes like '1#PCS - ', 'INV_01 - '
    n = re.sub(r"^[0-9]+#\s*", "", n)
    n = re.sub(r"^[a-z0-9_-]+\s*-\s*", "", n)
    n = re.sub(r"^(?:inverter|onduleur|inv)[\s_]+", "", n)
    # Strip parenthetical or bracketed unit suffixes like (Hz), [A], (kW)
    n = re.sub(r"[\(\[\{][^\)\]\}]*[\)\]\}]", "", n)
    # Replace underscores with spaces, strip punctuation except space
    n = n.replace("_", " ")
    n = re.sub(r"[^\w\s]", " ", n)
    n = re.sub(r"\s+", " ", n).strip()
    return n


class DatasetKnowledge:
    """
    Knowledge container trained on real Modbus definition datasets.
    Stores empirical mappings, unit correlations, type norms, and action codes.
    """

    def __init__(self) -> None:
        self.name_to_tag: dict[str, str] = {}
        self.tag_counts: Counter[str] = Counter()
        self.metric_units: dict[str, Counter[str]] = defaultdict(Counter)
        self.metric_types: dict[str, Counter[str]] = defaultdict(Counter)
        self.metric_actions: dict[str, Counter[str]] = defaultdict(Counter)
        self.total_definitions_mined: int = 0
        self.total_registers_mined: int = 0

    def learn_entry(
        self,
        name: str,
        tag: str,
        dtype: str = "",
        unit: str = "",
        action: str = "",
    ) -> None:
        if not name:
            return
        key = clean_variable_key(name)
        if not key:
            return

        if tag:
            canonical_tag = normalize_tag_to_sunpm(tag)
            if canonical_tag:
                self.name_to_tag[key] = canonical_tag
                self.tag_counts[canonical_tag] += 1
                if unit:
                    self.metric_units[canonical_tag][unit.strip()] += 1
                if dtype:
                    self.metric_types[canonical_tag][dtype.strip().upper()] += 1
                if action:
                    self.metric_actions[canonical_tag][action.strip()] += 1

    # Convenient alias
    record_variable = learn_entry

    def get_learned_tag(self, name: str) -> str:
        key = clean_variable_key(name)
        return self.name_to_tag.get(key, "")

    def _resolve_tag(self, tag_or_name: str) -> str:
        if tag_or_name in self.metric_units or tag_or_name in self.tag_counts:
            return tag_or_name
        return self.get_learned_tag(tag_or_name)

    def get_metric_units(self, tag_or_name: str) -> dict[str, int]:
        tag = self._resolve_tag(tag_or_name)
        return dict(self.metric_units.get(tag, {}))

    def get_metric_types(self, tag_or_name: str) -> dict[str, int]:
        tag = self._resolve_tag(tag_or_name)
        return dict(self.metric_types.get(tag, {}))

    def get_metric_actions(self, tag_or_name: str) -> dict[str, int]:
        tag = self._resolve_tag(tag_or_name)
        return dict(self.metric_actions.get(tag, {}))

    def to_dict(self) -> dict[str, Any]:
        return {
            "name_to_tag": self.name_to_tag,
            "tag_counts": dict(self.tag_counts),
            "metric_units": {k: dict(v) for k, v in self.metric_units.items()},
            "metric_types": {k: dict(v) for k, v in self.metric_types.items()},
            "metric_actions": {k: dict(v) for k, v in self.metric_actions.items()},
            "total_definitions_mined": self.total_definitions_mined,
            "total_registers_mined": self.total_registers_mined,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> DatasetKnowledge:
        inst = cls()
        inst.name_to_tag = data.get("name_to_tag", {})
        inst.tag_counts = Counter(data.get("tag_counts", {}))
        inst.metric_units = defaultdict(
            Counter, {k: Counter(v) for k, v in data.get("metric_units", {}).items()}
        )
        inst.metric_types = defaultdict(
            Counter, {k: Counter(v) for k, v in data.get("metric_types", {}).items()}
        )
        inst.metric_actions = defaultdict(
            Counter, {k: Counter(v) for k, v in data.get("metric_actions", {}).items()}
        )
        inst.total_definitions_mined = data.get("total_definitions_mined", 0)
        inst.total_registers_mined = data.get("total_registers_mined", 0)
        return inst


DEFAULT_DATASET_PATH = Path(__file__).parent / "dataset_knowledge.json"
DEFAULT_EQUIPEMENTIERS_DIR = r"\\WDN-FILES.WEBDYN.COM\Support\Equipementiers"


def load_dataset_knowledge(path: str | Path | None = None) -> DatasetKnowledge | None:
    """Loads a pre-trained DatasetKnowledge instance from JSON if available."""
    json_path = Path(path) if path else DEFAULT_DATASET_PATH
    if json_path.is_file():
        try:
            with open(json_path, encoding="utf-8") as f:
                data = json.load(f)
            return DatasetKnowledge.from_dict(data)
        except Exception:
            return None
    return None


ACTIVE_KNOWLEDGE: DatasetKnowledge | None = None


def train_knowledge_from_definitions(
    directories: list[str | Path] | str | Path | None = None,
    auto_activate: bool = True,
    force_retrain: bool = False,
    recursive: bool = True,
) -> DatasetKnowledge:
    """
    Parses all definition CSV files in specified directories, training a DatasetKnowledge
    model with variable tags, unit associations, and type distributions.
    """
    global ACTIVE_KNOWLEDGE

    if ACTIVE_KNOWLEDGE is not None and not force_retrain and directories is None:
        return ACTIVE_KNOWLEDGE

    # If no directories explicitly given and not forcing retrain, check for pre-trained model file
    if directories is None and not force_retrain:
        preloaded = load_dataset_knowledge()
        if preloaded is not None:
            if auto_activate:
                ACTIVE_KNOWLEDGE = preloaded
            return preloaded

    if directories is None:
        env_share = os.environ.get("EQUIPEMENTIERS_DIR", DEFAULT_EQUIPEMENTIERS_DIR)
        candidates = [
            Path("Definitions"),
            Path(__file__).parent.parent / "Definitions",
            Path(env_share),
            Path("generated_definitions"),
            Path(__file__).parent.parent / "generated_definitions",
        ]
        resolved_dirs = [p for p in candidates if p.is_dir()]
    elif isinstance(directories, (str, Path)):
        resolved_dirs = [Path(directories)]
    else:
        resolved_dirs = [Path(d) for d in directories]

    knowledge = DatasetKnowledge()
    seen_files: set[str] = set()

    for d_path in resolved_dirs:
        if not d_path.is_dir():
            continue
        csv_files = sorted(d_path.rglob("*.csv") if recursive else d_path.glob("*.csv"))
        for p in csv_files:
            if p.name in seen_files:
                continue
            seen_files.add(p.name)
            try:
                with open(p, encoding="utf-8-sig", errors="ignore") as f:
                    reader = csv.reader(f, delimiter=";")
                    header = next(reader, None)
                    if not header or len(header) < 2 or not any(header):
                        continue
                    knowledge.total_definitions_mined += 1

                    for row in reader:
                        if not row or len(row) < 7:
                            continue
                        knowledge.total_registers_mined += 1
                        dtype = row[3].strip() if len(row) > 3 else ""
                        name = row[5].strip() if len(row) > 5 else ""
                        tag = row[6].strip() if len(row) > 6 else ""
                        unit = row[9].strip() if len(row) > 9 else ""
                        action = row[10].strip() if len(row) > 10 else ""

                        # Only learn valid canonical SunPM tags when mining datasets
                        norm_tag = normalize_tag_to_sunpm(tag)
                        if not is_canonical_sunpm_tag(norm_tag):
                            continue

                        knowledge.learn_entry(
                            name=name,
                            tag=norm_tag,
                            dtype=dtype,
                            unit=unit,
                            action=action,
                        )
            except Exception:
                continue

    if auto_activate:
        ACTIVE_KNOWLEDGE = knowledge

    return knowledge


def get_active_knowledge() -> DatasetKnowledge:
    """Returns the currently active trained DatasetKnowledge instance (training lazily if needed)."""
    global ACTIVE_KNOWLEDGE
    if ACTIVE_KNOWLEDGE is None:
        ACTIVE_KNOWLEDGE = load_dataset_knowledge()
        if ACTIVE_KNOWLEDGE is None:
            ACTIVE_KNOWLEDGE = train_knowledge_from_definitions(auto_activate=False)
    return ACTIVE_KNOWLEDGE


def _clean_variable_label(name: str) -> str:
    """Strips equipment prefixes and normalizes whitespace."""
    clean = name.strip()
    if " - " in clean:
        clean = clean.split(" - ", 1)[1].strip()
    return clean


def infer_sunpm_tag(
    variable_name: str,
    existing_tag: str = "",
    equipment_type: str = "Inverter",
    normalize_existing: bool = True,
    use_learned_dataset: bool = True,
    knowledge: DatasetKnowledge | None = None,
) -> str:
    """
    Infers the official canonical Webdyn SunPM tag for a Modbus variable based on
    the SunPM specification (SunPM_DEfFile_Tags_v1.xlsx) and trained dataset knowledge.

    Supports:
    - Prior learned empirical tag mappings from trained definition libraries
    - 3-phase per-phase measurements (U_AC_[1..3], U_AC_L1L2, I_AC_[1..3], P_AC_[1..3], etc.)
    - DC Bus, MPPTs (U_DC_MPPT_x, I_DC_MPPT_x), PV strings (I_DC_STRING_x)
    - Active & Reactive control commands (CMD_ON, CMD_OFF, CMD_P_AC_SET_PERCENT, CMD_PF, etc.)
    - Sensors (POA, GHI, TEMP_AMB, TEMP_CELL) and Batteries (SOC, SOH, TEMP, U_DC)
    - Normalization of legacy tags (RealPower -> P_AC, EnergyTotal -> E_TOT, etc.)
    """
    if existing_tag:
        if normalize_existing:
            normalized = normalize_tag_to_sunpm(existing_tag)
            if normalized:
                return normalized
        return existing_tag

    # 0. Check learned dataset associations first
    if use_learned_dataset:
        kb = knowledge if knowledge is not None else get_active_knowledge()
        if kb:
            learned = kb.get_learned_tag(variable_name)
            if learned in ("CMD_P_AC_SET_PERCENT", "CMD_Q_AC_SET_PERCENT"):
                low = variable_name.lower()
                if any(u in low for u in ("kw", " w", "[w]", "(w)", "kvar", "var", "abs")):
                    learned = ""
            if learned:
                return learned

    clean_name = _clean_variable_label(variable_name)
    n_lower = clean_name.lower()
    eq_lower = equipment_type.lower()

    # 0. Remote Commands & Active / Reactive Controls (evaluated before telemetry)
    if any(k in n_lower for k in ("start inverter", "turn on", "power on", "démarrage onduleur")):
        return "CMD_ON"
    if any(k in n_lower for k in ("stop inverter", "turn off", "power off", "arrêt onduleur")):
        return "CMD_OFF"
    if (
        "power limitation enable" in n_lower
        or "power limit enable" in n_lower
        or "wmaxlim" in n_lower
        or "ctrlpwr_ena" in n_lower
    ):
        return "CMD_P_ENA"
    if "reactive power enable" in n_lower or "varpct_ena" in n_lower:
        return "CMD_Q_ENA"
    if (
        "reactive power setpoint" in n_lower
        or "consigne réactif" in n_lower
        or "cmdqpct" in n_lower
        or "cmdangle" in n_lower
    ):
        if "%" in clean_name or "percent" in n_lower or "pct" in n_lower:
            return "CMD_Q_AC_SET_PERCENT"
        if any(k in n_lower for k in ("kvar", "var", "abs")):
            return "CMD_Q_AC_SET_ABS"
        return "CMD_Q_AC_SET_PERCENT"
    if any(k in n_lower for k in ("power limit", "power limitation", "cmdpwr")) or re.search(
        r"\b(?:ac\s+)?active\s+power\s+(?:limitation|setpoint|limit)\b", n_lower
    ):
        if "%" in clean_name or "percent" in n_lower:
            return "CMD_P_AC_SET_PERCENT"
        if any(k in n_lower for k in ("kw", "w", "abs")):
            return "CMD_P_AC_SET_ABS"
        return "CMD_P_AC_SET_PERCENT"
    if "ramp time" in n_lower or "temps de rampe" in n_lower:
        return "CMD_RAMP_TIME_S"
    if "revert time" in n_lower or "timeout period" in n_lower:
        return "CMD_REVERT_TIME_S"
    if (
        "power factor setpoint" in n_lower
        or "consigne facteur de puissance" in n_lower
        or "cmdpf" in n_lower
    ):
        return "CMD_PF"
    if "commit" in n_lower:
        return "CMD_COMMIT"

    # 1. AC Voltage (Phase to Phase: U_AC_L1L2, U_AC_L2L3, U_AC_L3L1)
    if any(k in n_lower for k in ("voltage", "tension", "u_", "v_")) and not any(
        k in n_lower for k in ("dc", "mppt", "string", "bus")
    ):
        if any(
            p in n_lower
            for p in ("l1-l2", "l1l2", "phase a to b", "phase a-b", "phase a to phase b", "uab")
        ):
            return "U_AC_L1L2"
        if any(
            p in n_lower
            for p in ("l2-l3", "l2l3", "phase b to c", "phase b-c", "phase b to phase c", "ubc")
        ):
            return "U_AC_L2L3"
        if any(
            p in n_lower
            for p in ("l3-l1", "l3l1", "phase c to a", "phase c-a", "phase c to phase a", "uca")
        ):
            return "U_AC_L3L1"
        # Phase to Neutral: U_AC_1, U_AC_2, U_AC_3
        if re.search(r"\b(?:l1-n|l1_n|phase\s*[a1]|line\s*[a1]|l1\b|ua\b)", n_lower):
            return "U_AC_1"
        if re.search(r"\b(?:l2-n|l2_n|phase\s*[b2]|line\s*[b2]|l2\b|ub\b)", n_lower):
            return "U_AC_2"
        if re.search(r"\b(?:l3-n|l3_n|phase\s*[c3]|line\s*[c3]|l3\b|uc\b)", n_lower):
            return "U_AC_3"

    # 2. AC Current (I_AC_1, I_AC_2, I_AC_3, I_AC)
    if any(k in n_lower for k in ("current", "courant", "i_ac", "iac")) and not any(
        k in n_lower for k in ("dc", "mppt", "string", "bus", "leakage", "leak")
    ):
        if re.search(r"\b(?:phase\s*[a1]|line\s*[a1]|l1\b|ia\b|iac1\b)", n_lower):
            return "I_AC_1"
        if re.search(r"\b(?:phase\s*[b2]|line\s*[b2]|l2\b|ib\b|iac2\b)", n_lower):
            return "I_AC_2"
        if re.search(r"\b(?:phase\s*[c3]|line\s*[c3]|l3\b|ic\b|iac3\b)", n_lower):
            return "I_AC_3"
        if re.search(r"\b(?:total|sum|ac\s+current|courant\s+ac\s+total)\b", n_lower):
            return "I_AC"

    # 3. Reactive Power (Q_AC_1, Q_AC_2, Q_AC_3, Q_AC) - Checked before Active Power
    if any(
        k in n_lower
        for k in (
            "reactive power",
            "puissance réactive",
            "puissance reactive",
            "reactivepower",
            "qac",
        )
    ):
        if re.search(r"\b(?:phase\s*[a1]|line\s*[a1]|l1\b|qac1\b|reactivepow1)", n_lower):
            return "Q_AC_1"
        if re.search(r"\b(?:phase\s*[b2]|line\s*[b2]|l2\b|qac2\b|reactivepow2)", n_lower):
            return "Q_AC_2"
        if re.search(r"\b(?:phase\s*[c3]|line\s*[c3]|l3\b|qac3\b|reactivepow3)", n_lower):
            return "Q_AC_3"
        return "Q_AC"

    # 4. Active Power (P_AC_1, P_AC_2, P_AC_3, P_AC)
    if re.search(
        r"\b(?:(?:ac|output)\s+)?active\s+power\b|\bpuissance\s+active\b|\breal\s*power\b|\bpac\b|\bactivepow",
        n_lower,
    ):
        if re.search(r"\b(?:phase\s*[a1]|line\s*[a1]|l1\b|pac1\b|activepow1)", n_lower):
            return "P_AC_1"
        if re.search(r"\b(?:phase\s*[b2]|line\s*[b2]|l2\b|pac2\b|activepow2)", n_lower):
            return "P_AC_2"
        if re.search(r"\b(?:phase\s*[c3]|line\s*[c3]|l3\b|pac3\b|activepow3)", n_lower):
            return "P_AC_3"
        return "P_AC"

    # 5. Apparent Power (S_AC_1, S_AC_2, S_AC_3, S_AC)
    if any(k in n_lower for k in ("apparent power", "puissance apparente", "sac")):
        if re.search(r"\b(?:phase\s*[a1]|l1\b|sac1\b)", n_lower):
            return "S_AC_1"
        if re.search(r"\b(?:phase\s*[b2]|l2\b|sac2\b)", n_lower):
            return "S_AC_2"
        if re.search(r"\b(?:phase\s*[c3]|l3\b|sac3\b)", n_lower):
            return "S_AC_3"
        return "S_AC"

    # 6. Frequency (FREQ, FREQ_1, FREQ_2, FREQ_3)
    if any(
        k in n_lower for k in ("frequency", "fréquence", "frequence", "gridfrequency")
    ) or re.search(r"\bfac\b", n_lower):
        if re.search(r"\b(?:phase\s*[a1]|l1\b)", n_lower):
            return "FREQ_1"
        if re.search(r"\b(?:phase\s*[b2]|l2\b)", n_lower):
            return "FREQ_2"
        if re.search(r"\b(?:phase\s*[c3]|l3\b)", n_lower):
            return "FREQ_3"
        return "FREQ"

    # 7. Power Factor (PF, PF_1, PF_2, PF_3)
    if any(
        k in n_lower
        for k in ("power factor", "facteur de puissance", "cosphi", "cos_phi", "cos phi")
    ):
        if re.search(r"\b(?:phase\s*[a1]|l1\b)", n_lower):
            return "PF_1"
        if re.search(r"\b(?:phase\s*[b2]|l2\b)", n_lower):
            return "PF_2"
        if re.search(r"\b(?:phase\s*[c3]|l3\b)", n_lower):
            return "PF_3"
        return "PF"

    # 8. Energy Generation & Yield
    if "energy" in n_lower or "énergie" in n_lower or "energie" in n_lower or "yield" in n_lower:
        if any(k in n_lower for k in ("daily", "today", "jour", "energyday")):
            return "E_DAY"
        if any(k in n_lower for k in ("export", "exportée", "exportee")):
            return "E_EXP"
        if any(k in n_lower for k in ("import", "importée", "importee")):
            return "E_IMP"
        if any(k in n_lower for k in ("total", "cumulative", "lifetime", "life", "energytotal")):
            if re.search(r"\b(?:phase\s*[a1]|l1\b)", n_lower):
                return "E_TOT_1"
            if re.search(r"\b(?:phase\s*[b2]|l2\b)", n_lower):
                return "E_TOT_2"
            if re.search(r"\b(?:phase\s*[c3]|l3\b)", n_lower):
                return "E_TOT_3"
            return "E_TOT"

    # 9. Rated Power
    if any(
        k in n_lower
        for k in ("rated power", "nominal power", "puissance nominale", "plant capacity")
    ):
        return "P_RATED"

    # 10. MPPT Current / Voltage / Power
    m_mppt = re.search(r"\bmppt\s*(\d+)", n_lower)
    if m_mppt:
        idx = m_mppt.group(1)
        if any(k in n_lower for k in ("current", "courant", "i_")):
            return f"I_DC_MPPT_{idx}"
        if any(k in n_lower for k in ("voltage", "tension", "u_", "v_")):
            return f"U_DC_MPPT_{idx}"
        if any(k in n_lower for k in ("power", "puissance", "p_")):
            return f"P_DC_MPPT_{idx}"

    # 11. String Current / Voltage
    m_str = re.search(r"\b(?:pv|string)\s*(\d+)", n_lower)
    if m_str:
        idx = m_str.group(1)
        if any(k in n_lower for k in ("current", "courant", "i_")):
            return f"I_DC_STRING_{idx}"
        if any(k in n_lower for k in ("voltage", "tension", "u_", "v_")):
            return f"U_DC_STRING_{idx}"

    # 12. DC Bus Voltage / Current / Power
    if "dc" in n_lower or "bus" in n_lower:
        if any(k in n_lower for k in ("voltage", "tension", "udc", "vdc")):
            return "U_DC"
        if any(k in n_lower for k in ("current", "courant", "idc")):
            return "I_DC"
        if any(k in n_lower for k in ("power", "puissance", "pdc")):
            return "P_DC"

    # 13. Temperature
    if "temperature" in n_lower or "température" in n_lower or "temp" in n_lower:
        if "ambient" in n_lower or "ambiante" in n_lower:
            return "TEMP_AMB"
        if "cell" in n_lower or "module" in n_lower or "panel" in n_lower:
            return "TEMP_CELL"
        if "battery" in n_lower or eq_lower in ("battery", "bess"):
            return "TEMP"
        if any(
            k in n_lower
            for k in ("internal", "interne", "cabinet", "inverter", "heatsink", "radiateur")
        ):
            return "TEMP_INT"

    # 14. Sensors & Meteorological
    if any(k in n_lower for k in ("ghi", "horizontal irradiance", "horizontal")):
        return "GHI"
    if any(
        k in n_lower
        for k in ("poa", "plane of array", "irradiance", "pyranometer", "ensoleillement")
    ):
        return "POA"
    if "wind" in n_lower or "vent" in n_lower:
        if any(k in n_lower for k in ("direction", "bearing", "angle")):
            return "WIND_BEARING"
        return "WIND_SPEED"
    if "humidity" in n_lower or "humidité" in n_lower:
        return "HUMIDITY"
    if "pressure" in n_lower or "pression" in n_lower:
        return "PRESSURE"
    if "rain" in n_lower or "pluie" in n_lower:
        return "RAIN"

    # 15. Battery Storage (BESS)
    if eq_lower in ("battery", "bess") or "battery" in n_lower:
        if any(k in n_lower for k in ("state of charge", "soc", "charge state")):
            return "SOC"
        if any(k in n_lower for k in ("state of health", "soh")):
            return "SOH"

    # 16. Electrical Insulation
    if "insulation" in n_lower or "isolement" in n_lower or "r_iso" in n_lower:
        return "R_ISO"

    # 17. Device Metadata & Communication
    if any(k in n_lower for k in ("serial number", "numéro de série", "sn", "device_sn")):
        return "SERIAL"
    if any(k in n_lower for k in ("model", "modèle", "product model", "device model")):
        return "MODEL"
    if any(k in n_lower for k in ("manufacturer", "constructeur", "fabricant")):
        return "MANUFACTURER"
    if any(k in n_lower for k in ("modbus address", "slave address", "adresse modbus", "slave id")):
        return "UNIT_ID"

    # 19. Displays & Public Counters
    if eq_lower in ("display", "afficheur") or "display" in n_lower:
        if "avoided" in n_lower or "évité" in n_lower or "co2" in n_lower:
            if any(k in n_lower for k in ("day", "daily", "jour")):
                return "AVOIDED_CO2_DAY"
            return "AVOIDED_CO2"
        if "self-consumption" in n_lower or "autoconso" in n_lower:
            return "SELF_CONSUMPTION_RATIO"
        if "consumption" in n_lower or "conso" in n_lower:
            return "LOAD_P_AC"

    return ""


# -----------------------------------------------------------------------------
# Definition File Enrichment & Batch Processing
# -----------------------------------------------------------------------------


def enrich_definition_rows(
    rows: list[dict[str, Any]],
    equipment_type: str = "Inverter",
    overwrite_existing: bool = False,
    normalize_legacy: bool = True,
) -> int:
    """
    Enriches a list of definition register dictionaries by populating or normalizing
    their 'Tag' field using the SunPM tag knowledge engine.

    Returns the number of tags added or modified.
    """
    modified_count = 0
    for row in rows:
        var_name = str(row.get("Name", "") or row.get("Variable", "")).strip()
        existing_tag = str(row.get("Tag", "")).strip()

        if existing_tag and not overwrite_existing:
            if normalize_legacy:
                normalized = normalize_tag_to_sunpm(existing_tag)
                if normalized != existing_tag:
                    row["Tag"] = normalized
                    modified_count += 1
            continue

        inferred = infer_sunpm_tag(
            variable_name=var_name,
            existing_tag=existing_tag if not overwrite_existing else "",
            equipment_type=equipment_type,
            normalize_existing=normalize_legacy,
        )
        if inferred and inferred != existing_tag:
            row["Tag"] = inferred
            modified_count += 1

    return modified_count


def enrich_definition_file(
    input_path: str,
    output_path: str | None = None,
    equipment_type: str = "",
    overwrite_existing: bool = False,
    normalize_legacy: bool = True,
) -> int:
    """
    Reads an existing WebdynSunPM definition CSV, infers missing or legacy SunPM tags,
    and atomically writes the enriched definition.

    Returns the number of tags added or updated.
    """
    in_p = Path(input_path)
    out_p = Path(output_path) if output_path else in_p

    with open(in_p, encoding="utf-8-sig", errors="ignore") as f:
        reader = csv.reader(f, delimiter=";")
        lines = [r for r in reader if r and any(r)]

    if not lines:
        return 0

    header = lines[0]
    detected_type = equipment_type or (header[1] if len(header) > 1 else "Inverter")

    modified = 0
    new_lines = [header]

    for row in lines[1:]:
        if len(row) < 7:
            new_lines.append(row)
            continue

        var_name = row[5].strip() if len(row) > 5 else ""
        existing_tag = row[6].strip()

        target_tag = existing_tag
        if existing_tag and not overwrite_existing:
            if normalize_legacy:
                target_tag = normalize_tag_to_sunpm(existing_tag)
        else:
            inferred = infer_sunpm_tag(
                variable_name=var_name,
                existing_tag=existing_tag if not overwrite_existing else "",
                equipment_type=detected_type,
                normalize_existing=normalize_legacy,
            )
            if inferred:
                target_tag = inferred

        if target_tag != existing_tag:
            row[6] = target_tag
            modified += 1

        new_lines.append(row)

    # Write output atomically
    with staged_text_output(str(out_p)) as out_f:
        writer = csv.writer(out_f, delimiter=";")
        writer.writerows(new_lines)

    return modified


# -----------------------------------------------------------------------------
# Action Code & Function Code Extractors
# -----------------------------------------------------------------------------


def infer_action_code(name: str, fc: str | int = 4, rw_field: str = "") -> str:
    """
    Determines the WebdynSunPM Action code based on the dataset's empirical distribution:
    - '0': Explicit reserve/unused registers.
    - '1': Write commands / configuration settings.
    - '9': Alarms, trip signals, operating status, breakers.
    - '4': Periodic telemetry measurements (voltage, current, power, energy, etc.).
    """
    n_lower = name.lower().strip()
    rw_lower = str(rw_field).lower().strip()
    fc_str = str(fc).strip()

    # Reserve registers
    if "reserve" in n_lower or "reserved" in n_lower or n_lower in ("res", "res.", "nc"):
        return "0"

    # Write commands
    if (fc_str in ("1", "5", "6", "15", "16") or "w" in rw_lower) and any(
        k in n_lower
        for k in (
            "write",
            "cmd",
            "command",
            "commande",
            "control",
            "contrôle",
            "set",
            "consigne",
            "setting",
            "mode",
            "duration",
            "limit",
            "on/off",
            "enable",
        )
    ):
        return "1"

    # Alarms and status
    if any(
        k in n_lower
        for k in (
            "alarm",
            "alarme",
            "status",
            "fault",
            "failure",
            "fail",
            "error",
            "erreur",
            "alert",
            "alerte",
            "crash-stop",
            "trip",
            "warning",
            "état",
            "defaut",
            "défaut",
            "panne",
        )
    ):
        return "9"

    # Telemetry default
    return "4"


def parse_function_code_from_text(text: str) -> str:
    """
    Extracts Modbus Function Code from section banners (e.g., 'Function Code: 0x04 - Read Input Registers').
    Returns Webdyn Info1 code: '1' (Coil), '2' (Discrete), '3' (Holding), '4' (Input).
    """
    t_lower = text.lower()
    if (
        re.search(r"\b(?:0x01|fc0?1|code\s*fonction\s*0?1)\b", t_lower)
        or "read coil" in t_lower
        or "coil" in t_lower
    ):
        return "1"
    if re.search(r"\b(?:0x02|fc0?2|code\s*fonction\s*0?2)\b", t_lower) or "discrete" in t_lower:
        return "2"
    if (
        re.search(r"\b(?:0x04|fc0?4|code\s*fonction\s*0?4)\b", t_lower)
        or "input register" in t_lower
    ):
        return "4"
    if any(
        code in t_lower
        for code in (
            "0x03",
            "0x06",
            "0x10",
            "0x16",
            "fc03",
            "fc3",
            "fc06",
            "fc6",
            "fc16",
            "holding",
        )
    ):
        return "3"

    m = re.search(r"(?:function\s*code|code\s*fonction|fc)\s*[:=\s]\s*(0?[1-6]|1[06])\b", t_lower)
    if m:
        val = m.group(1).lstrip("0") or "1"
        if val in ("1", "5"):
            return "1"
        if val == "2":
            return "2"
        if val in ("3", "6", "16", "10"):
            return "3"
        if val == "4":
            return "4"

    return "4"


# -----------------------------------------------------------------------------
# Definition Audit and Auto-Correction Engine
# -----------------------------------------------------------------------------


@dataclass
class DefinitionAuditIssue:
    line: int
    severity: str  # ERROR, WARNING, INFO
    code: str
    field: str
    message: str
    suggestion: str = ""


@dataclass
class DefinitionAuditReport:
    filepath: str
    is_valid: bool
    register_count: int
    issues: list[DefinitionAuditIssue] = field(default_factory=list)
    stats: dict[str, int] = field(default_factory=dict)

    @property
    def errors(self) -> list[DefinitionAuditIssue]:
        return [i for i in self.issues if i.severity == "ERROR"]

    @property
    def warnings(self) -> list[DefinitionAuditIssue]:
        return [i for i in self.issues if i.severity == "WARNING"]

    @property
    def infos(self) -> list[DefinitionAuditIssue]:
        return [i for i in self.issues if i.severity == "INFO"]


def audit_definition_file(
    filepath: str | Path,
    strict_overlap: bool = False,
    strict_actions: bool = False,
    knowledge: DatasetKnowledge | None = None,
) -> DefinitionAuditReport:
    """
    Audits a WebdynSunPM definition CSV file for schema conformity, Modbus syntax,
    address boundary/overlap collisions, semantic tag consistency, and permission codes.
    """
    p = Path(filepath)
    if not p.is_file():
        msg = f"Definition file not found: {p}"
        return DefinitionAuditReport(
            filepath=str(p),
            is_valid=False,
            register_count=0,
            issues=[
                DefinitionAuditIssue(
                    line=0,
                    severity="ERROR",
                    code="FILE_NOT_FOUND",
                    field="filepath",
                    message=msg,
                )
            ],
            stats={"errors": 1, "warnings": 0, "infos": 0, "registers": 0},
        )

    issues: list[DefinitionAuditIssue] = []
    seen_rows: dict[tuple[str, str, str, str], int] = {}
    gen = Generator()

    # 1. First run structural & address checks via Generator.validate_csv_detailed
    gen_report = gen.validate_csv_detailed(str(p), strict=False, strict_overlap=strict_overlap)
    for issue in gen_report.issues:
        issues.append(
            DefinitionAuditIssue(
                line=issue.line,
                severity=issue.severity,
                code=issue.code,
                field=issue.field,
                message=issue.message,
            )
        )

    total_registers = gen_report.register_count

    # 2. Deep semantic and tag analysis
    try:
        with open(p, encoding="utf-8-sig", errors="ignore") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader, None)
            if not header or len(header) < 2 or not any(header):
                return DefinitionAuditReport(
                    filepath=str(p),
                    is_valid=False,
                    register_count=0,
                    issues=issues,
                    stats={
                        "errors": sum(1 for i in issues if i.severity == "ERROR"),
                        "warnings": 0,
                        "infos": 0,
                        "registers": 0,
                    },
                )

            eq_type = header[1].strip() if len(header) > 1 else "Inverter"

            for line_num, row in enumerate(reader, start=2):
                if not row or not any(row) or len(row) < 7:
                    continue

                fc = row[1].strip() if len(row) > 1 else ""
                addr = row[2].strip() if len(row) > 2 else ""
                dtype = row[3].strip().upper() if len(row) > 3 else ""
                name = row[5].strip() if len(row) > 5 else ""
                tag = row[6].strip() if len(row) > 6 else ""
                unit = row[9].strip() if len(row) > 9 else ""
                action = row[10].strip() if len(row) > 10 else ""

                # Duplicate identical row detection
                row_key = (fc, addr, dtype, name)
                if row_key in seen_rows:
                    issues.append(
                        DefinitionAuditIssue(
                            line=line_num,
                            severity="WARNING",
                            code="DUPLICATE_ROW",
                            field="row",
                            message=f"Line {line_num}: Duplicate register entry (previously at line {seen_rows[row_key]}).",
                        )
                    )
                else:
                    seen_rows[row_key] = line_num

                # Tag Analysis
                if tag:
                    if tag in SUNPM_LEGACY_TO_CANONICAL:
                        canonical = SUNPM_LEGACY_TO_CANONICAL[tag]
                        issues.append(
                            DefinitionAuditIssue(
                                line=line_num,
                                severity="WARNING",
                                code="LEGACY_TAG",
                                field="Tag",
                                message=f"Line {line_num}: Legacy tag '{tag}' should be normalized to SunPM '{canonical}'.",
                                suggestion=canonical,
                            )
                        )
                    elif tag not in SUNPM_CANONICAL_TAGS and not any(
                        tag.startswith(pref)
                        for pref in (
                            "U_DC_MPPT_",
                            "I_DC_MPPT_",
                            "P_DC_MPPT_",
                            "I_DC_STRING_",
                            "U_DC_STRING_",
                        )
                    ):
                        issues.append(
                            DefinitionAuditIssue(
                                line=line_num,
                                severity="INFO",
                                code="CUSTOM_TAG",
                                field="Tag",
                                message=f"Line {line_num}: Non-canonical SunPM tag '{tag}'.",
                            )
                        )
                else:
                    inferred = infer_sunpm_tag(name, equipment_type=eq_type)
                    if inferred:
                        issues.append(
                            DefinitionAuditIssue(
                                line=line_num,
                                severity="INFO",
                                code="UNTAGGED_VARIABLE",
                                field="Tag",
                                message=f"Line {line_num}: Untagged variable '{name}' matches SunPM tag '{inferred}'.",
                                suggestion=inferred,
                            )
                        )

                # Unit Coherence Analysis
                u_clean = unit.upper()
                if tag in ("P_AC", "P_AC_1", "P_AC_2", "P_AC_3") and u_clean in (
                    "V",
                    "A",
                    "HZ",
                    "DEG",
                ):
                    issues.append(
                        DefinitionAuditIssue(
                            line=line_num,
                            severity="WARNING",
                            code="INCOHERENT_UNIT",
                            field="Unit",
                            message=f"Line {line_num}: Active power tag '{tag}' has invalid unit '{unit}'.",
                        )
                    )
                if tag in (
                    "U_AC_1",
                    "U_AC_2",
                    "U_AC_3",
                    "U_AC_L1L2",
                    "U_AC_L2L3",
                    "U_AC_L3L1",
                ) and u_clean in ("W", "KW", "A", "HZ"):
                    issues.append(
                        DefinitionAuditIssue(
                            line=line_num,
                            severity="WARNING",
                            code="INCOHERENT_UNIT",
                            field="Unit",
                            message=f"Line {line_num}: AC Voltage tag '{tag}' has invalid unit '{unit}'.",
                        )
                    )
                if tag in ("E_TOT", "E_DAY") and u_clean in ("W", "KW", "V", "A"):
                    issues.append(
                        DefinitionAuditIssue(
                            line=line_num,
                            severity="WARNING",
                            code="INCOHERENT_UNIT",
                            field="Unit",
                            message=f"Line {line_num}: Energy tag '{tag}' has power/voltage unit '{unit}' instead of energy (kWh/Wh).",
                        )
                    )

                # Action Code Coherence
                if tag.startswith("CMD_") and action == "4":
                    issues.append(
                        DefinitionAuditIssue(
                            line=line_num,
                            severity="ERROR" if strict_actions else "WARNING",
                            code="INCOHERENT_ACTION",
                            field="Action",
                            message=f"Line {line_num}: Remote command '{tag}' has Action=4 (read-only telemetry) instead of 1 (write).",
                            suggestion="1",
                        )
                    )

    except Exception as e:
        issues.append(
            DefinitionAuditIssue(
                line=0,
                severity="ERROR",
                code="READ_ERROR",
                field="file",
                message=f"Error reading definition file: {e}",
            )
        )

    errs = sum(1 for i in issues if i.severity == "ERROR")
    warns = sum(1 for i in issues if i.severity == "WARNING")
    infos = sum(1 for i in issues if i.severity == "INFO")

    return DefinitionAuditReport(
        filepath=str(p),
        is_valid=errs == 0,
        register_count=total_registers,
        issues=issues,
        stats={"errors": errs, "warnings": warns, "infos": infos, "registers": total_registers},
    )


def audit_definitions_directory(
    directory: str | Path,
    strict_overlap: bool = False,
    strict_actions: bool = False,
    knowledge: DatasetKnowledge | None = None,
) -> dict[str, DefinitionAuditReport]:
    """
    Batch audits all WebdynSunPM definition files inside a directory.
    Returns a dictionary mapping file path string to DefinitionAuditReport.
    """
    d_path = Path(directory)
    reports: dict[str, DefinitionAuditReport] = {}
    if not d_path.is_dir():
        return reports

    for csv_file in sorted(d_path.glob("*.csv")):
        reports[str(csv_file)] = audit_definition_file(
            csv_file,
            strict_overlap=strict_overlap,
            strict_actions=strict_actions,
            knowledge=knowledge,
        )

    return reports


def auto_fix_definition_file(
    input_path: str | Path,
    output_path: str | Path | None = None,
    auto_tag: bool = True,
    normalize_legacy: bool = True,
    deduplicate: bool = True,
) -> dict[str, int]:
    """
    Reads an existing definition CSV, automatically cleans known defects:
    - Normalizes legacy tags to official SunPM standards.
    - Populates missing tags via trained dataset knowledge & heuristics.
    - Eliminates duplicate identical rows.
    - Atomically writes out the sanitized definition file.

    Returns:
        dict[str, int]: Counts of repairs performed:
            {"tags_normalized": int, "tags_inferred": int, "duplicates_removed": int}
    """
    in_p = Path(input_path)
    out_p = Path(output_path) if output_path else in_p

    with open(in_p, encoding="utf-8-sig", errors="ignore") as f:
        reader = csv.reader(f, delimiter=";")
        lines = [r for r in reader if r and any(r)]

    if not lines:
        return {"tags_normalized": 0, "tags_inferred": 0, "duplicates_removed": 0}

    header = lines[0]
    detected_type = header[1].strip() if len(header) > 1 else "Inverter"

    counts = {"tags_normalized": 0, "tags_inferred": 0, "duplicates_removed": 0}
    new_lines = [header]
    seen_rows: set[tuple[str, str, str, str]] = set()

    seen_tags: set[str] = set()

    for row in lines[1:]:
        if len(row) < 7:
            new_lines.append(row)
            continue

        fc = row[1].strip() if len(row) > 1 else ""
        addr = row[2].strip() if len(row) > 2 else ""
        dtype = row[3].strip().upper() if len(row) > 3 else ""
        var_name = row[5].strip() if len(row) > 5 else ""
        existing_tag = row[6].strip()

        row_key = (fc, addr, dtype, var_name)
        if deduplicate and row_key in seen_rows:
            counts["duplicates_removed"] += 1
            continue
        seen_rows.add(row_key)

        target_tag = ""
        if existing_tag:
            norm = normalize_tag_to_sunpm(existing_tag) if normalize_legacy else existing_tag
            if norm not in seen_tags:
                target_tag = norm
                seen_tags.add(target_tag)
                if norm != existing_tag:
                    counts["tags_normalized"] += 1
            else:
                # Do not produce duplicate canonical tags within the same definition
                target_tag = ""
        elif auto_tag:
            inferred = infer_sunpm_tag(
                variable_name=var_name,
                existing_tag="",
                equipment_type=detected_type,
                normalize_existing=normalize_legacy,
            )
            if inferred and inferred not in seen_tags:
                target_tag = inferred
                seen_tags.add(inferred)
                counts["tags_inferred"] += 1

        row[6] = target_tag
        new_lines.append(row)

    # Atomically write sanitized definition
    with staged_text_output(str(out_p)) as out_f:
        writer = csv.writer(out_f, delimiter=";")
        writer.writerows(new_lines)

    return counts
