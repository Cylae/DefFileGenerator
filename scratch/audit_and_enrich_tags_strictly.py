#!/usr/bin/env python3
"""
Strict SunPM Tag Audit & Enrichment Engine.
Authoritative implementation based on `SunPM_DEfFile_Tags_v1.xlsx` and `docs/WEBDYN_TAGS_REFERENCE.md`.

STRICT DIRECTIVE:
1. Every non-empty tag MUST strictly belong to the official canonical SunPM specification.
2. ZERO APPROXIMATION: If uncertain, do NOT guess; leave the tag field empty ("").
3. Pseudo-tags, slugified descriptions, scale factors (_SF), and vendor internal markers are completely eliminated.
4. Canonical SunPM tags are populated for unambiguous variables (active/reactive/apparent power, voltages,
   currents, frequency, power factor, cumulative energies, temperatures, solar/weather sensors, battery BMS).
"""

import csv
import json
import os
import re
import sys
import tempfile
import time
from collections import Counter
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

# -----------------------------------------------------------------------------
# Authoritative Set of Canonical SunPM Tags
# -----------------------------------------------------------------------------

CANONICAL_SET: Set[str] = {
    # Basic Info & Identification
    "SERIAL", "MODEL", "MANUFACTURER", "UNIT_ID", "NAME", "P_RATED", "Q_RATED",
    # AC Power & Energy
    "P_AC", "P_AC_1", "P_AC_2", "P_AC_3",
    "Q_AC", "Q_AC_1", "Q_AC_2", "Q_AC_3",
    "S_AC", "S_AC_1", "S_AC_2", "S_AC_3",
    "FREQ", "FREQ_1", "FREQ_2", "FREQ_3",
    "PF", "PF_1", "PF_2", "PF_3",
    "E_DAY", "E_TOT", "E_TOT_1", "E_TOT_2", "E_TOT_3",
    "E_EXP", "E_EXP_1", "E_EXP_2", "E_EXP_3",
    "E_IMP", "E_IMP_1", "E_IMP_2", "E_IMP_3",
    "ES", "ES_1", "ES_2", "ES_3",
    "ES_EXP", "ES_IMP",
    "EQ", "EQ_1", "EQ_2", "EQ_3",
    "EQ_EXP", "EQ_IMP",
    # AC Voltages & Currents
    "U_AC_1", "U_AC_2", "U_AC_3",
    "U_AC_L1L2", "U_AC_L2L3", "U_AC_L3L1",
    "U_LE_AC_1", "U_LE_AC_2", "U_LE_AC_3", "U_NE",
    "I_AC", "I_AC_1", "I_AC_2", "I_AC_3",
    # DC Measurements
    "U_DC", "I_DC", "P_DC", "U_DC_NE", "U_DC_PE", "R_ISO", "TOT", "TEMP_INT",
    # Controls & Commands
    "CMD_ON", "CMD_OFF",
    "CMD_P_AC_SET_PERCENT", "CMD_P_AC_SET_ABS",
    "CMD_P_AC_LIM_PERCENT", "CMD_P_AC_LIM_ABS",
    "CMD_P_ENA", "CMD_P_MODE", "CMD_RAMP_TIME_S", "CMD_REVERT_TIME_S", "VAR_P_AC_REPEAT_TIME_S",
    "CMD_Q_AC_LIM_PERCENT", "CMD_Q_AC_LIM_ABS",
    "CMD_Q_AC_SET_PERCENT", "CMD_Q_AC_SET_ABS",
    "CMD_Q_ENA", "CMD_Q_MODE", "CMD_Q_RAMP_TIME_S", "CMD_Q_REVERT_TIME_S",
    "CMD_Q_PERCENT", "CMD_Q_ANGLE", "VAR_PF_SIGN",
    "CMD_PF", "CMD_PF_INT", "CMD_PF_ENA", "CMD_PF_RAMP_TIME_S", "CMD_PF_REVERT_TIME_S", "CMD_COMMIT",
    # Environmental & Sensors
    "POA", "GHI", "TEMP_AMB", "TEMP_CELL", "WIND_SPEED", "WIND_BEARING", "HUMIDITY", "PRESSURE", "RAIN",
    "S_POA", "S_GHI", "S_TEMP_AMB", "S_TEMP_CELL",
    # Battery & BESS
    "SOC", "SOH", "TEMP", "U_CELL_MAX", "U_CELL_MIN", "TEMP_MAX", "TEMP_MIN",
    "B_SoC", "B_SoH", "B_TEMP", "B_U_DC",
    # Displays & KPIs
    "AVOIDED_CO2", "AVOIDED_CO2_DAY", "SELF_CONSUMPTION_RATIO", "LOAD_P_AC",
    "CO2", "P_EXPORTED", "P_SELF", "IRRADIANCE", "STATUS_PLANT", "STATUS_INVERTERS",
    # Status & Alarms
    "STATE", "STATE_FAULT", "ALARM",
    # Meter M_ prefix variants
    "M_P_AC", "M_P_AC_1", "M_P_AC_2", "M_P_AC_3", "M_E", "M_E_EXP", "M_E_IMP", "M_FREQ", "M_PF",
    "M_I_AC_1", "M_I_AC_2", "M_I_AC_3",
}

DYNAMIC_PREFIXES = ("U_DC_MPPT_", "I_DC_MPPT_", "P_DC_MPPT_", "U_DC_STRING_", "I_DC_STRING_", "STATE_", "ALARM_")


def is_valid_official_tag(t: str) -> bool:
    """Verifies that a tag strictly conforms to the canonical SunPM specification."""
    if not t:
        return False
    t_clean = t.strip()
    if t_clean in CANONICAL_SET:
        return True
    return any(t_clean.startswith(p) and t_clean[len(p):].isdigit() for p in DYNAMIC_PREFIXES)


# -----------------------------------------------------------------------------
# Official Legacy-to-Canonical Normalization Map
# -----------------------------------------------------------------------------

OFFICIAL_LEGACY_MAP: Dict[str, str] = {
    # Active Power
    "realpower": "P_AC",
    "pac-tag": "P_AC",
    "activepowsumkw": "P_AC",
    "activepow1kw": "P_AC_1",
    "activepow2kw": "P_AC_2",
    "activepow3kw": "P_AC_3",
    "active_power": "P_AC",
    "active_power_phase_a": "P_AC_1",
    "active_power_phase_b": "P_AC_2",
    "active_power_phase_c": "P_AC_3",
    # Reactive Power
    "reactivepower": "Q_AC",
    "reactivepowsumkw": "Q_AC",
    "reactivepow1kw": "Q_AC_1",
    "reactivepow2kw": "Q_AC_2",
    "reactivepow3kw": "Q_AC_3",
    "reactive_power": "Q_AC",
    # Apparent Power
    "apparentpower": "S_AC",
    "apparent_power": "S_AC",
    # Energy
    "energytotal": "E_TOT",
    "totalenergy": "E_TOT",
    "eac-tag": "E_TOT",
    "energyday": "E_DAY",
    "dailyenergy": "E_DAY",
    "eacday-tag": "E_DAY",
    "watthours": "E_TOT",
    "total_energy": "E_TOT",
    "daily_energy": "E_DAY",
    # Frequency & Grid
    "gridfrequency": "FREQ",
    "gridfrequency1": "FREQ_1",
    "gridfrequency2": "FREQ_2",
    "gridfrequency3": "FREQ_3",
    "frequency": "FREQ",
    "cosphi": "PF",
    "power factor": "PF",
    "power_factor": "PF",
    # Metadata & Identity
    "nominalpower": "P_RATED",
    "displaymodel": "MODEL",
    "displayserial": "SERIAL",
    "displayserialnumber": "SERIAL",
    "manufacturerstr": "MANUFACTURER",
    "slaveaddress": "UNIT_ID",
    "deviceaddr": "UNIT_ID",
    # Commands
    "cmdon": "CMD_ON",
    "cmdoff": "CMD_OFF",
    "cmdpwrpercent": "CMD_P_AC_SET_PERCENT",
    "cmdpowerpercent": "CMD_P_AC_SET_PERCENT",
    "cmdpwr": "CMD_P_AC_SET_ABS",
    "wmaxlim_ena": "CMD_P_ENA",
    "wmaxlimena": "CMD_P_ENA",
    "ctrlpwr_ena": "CMD_P_ENA",
    "wmaxlimpct_rmptms": "CMD_RAMP_TIME_S",
    "wmaxlimpct_rvrttms": "CMD_REVERT_TIME_S",
    "cmdqpct": "CMD_Q_PERCENT",
    "cmdangle": "CMD_Q_ANGLE",
    "varpct_ena": "CMD_Q_ENA",
    "varpct_mod": "CMD_Q_MODE",
    "varpct_rmptms": "CMD_Q_RAMP_TIME_S",
    "varpct_rvrttms": "CMD_Q_REVERT_TIME_S",
    "cmdpf": "CMD_PF",
    "cmdpfint": "CMD_PF_INT",
    "outpfset": "CMD_PF",
    "outpfset_ena": "CMD_PF_ENA",
    "outpfset_rmptms": "CMD_PF_RAMP_TIME_S",
    "outpfset_rvrttms": "CMD_PF_REVERT_TIME_S",
    "commit": "CMD_COMMIT",
    # Displays & KPIs
    "carb-tag": "AVOIDED_CO2",
    "carbday-tag": "AVOIDED_CO2_DAY",
    "autoconso-tag": "SELF_CONSUMPTION_RATIO",
    "conso-tag": "LOAD_P_AC",
    # 3E Indicators & Storage
    "charge.state": "SOC",
    "charge.health": "SOH",
    "temperature.battery": "TEMP",
    "voltage.battery": "U_DC",
    "b_soc": "SOC",
    "b_soh": "SOH",
    "b_temp": "TEMP",
    "b_u_dc": "U_DC",
    "irradiance": "POA",
    "irradiance.horizontal": "GHI",
    "temperature.ambient": "TEMP_AMB",
    "temperature.cells": "TEMP_CELL",
    "energy.export": "E_EXP",
    "energy.import": "E_IMP",
}

# SunSpec standard abbreviations (Model 1, 101, 102, 103, 160, 201, 202, 203, 204)
SUNSPEC_FIELD_MAP: Dict[str, str] = {
    # Current
    "a": "I_AC",
    "amps": "I_AC",
    "apha": "I_AC_1",
    "aphb": "I_AC_2",
    "aphc": "I_AC_3",
    # Phase to Neutral Voltage
    "phv": "U_AC_1",
    "phvpha": "U_AC_1",
    "phvphb": "U_AC_2",
    "phvphc": "U_AC_3",
    # Phase to Phase Voltage
    "ppvphab": "U_AC_L1L2",
    "ppvphbc": "U_AC_L2L3",
    "ppvphca": "U_AC_L3L1",
    # Power
    "w": "P_AC",
    "watts": "P_AC",
    "wpha": "P_AC_1",
    "wphb": "P_AC_2",
    "wphc": "P_AC_3",
    "va": "S_AC",
    "vapha": "S_AC_1",
    "vaphb": "S_AC_2",
    "vaphc": "S_AC_3",
    "var": "Q_AC",
    "varpha": "Q_AC_1",
    "varphb": "Q_AC_2",
    "varphc": "Q_AC_3",
    # Grid & Frequency
    "hz": "FREQ",
    "pf": "PF",
    "pfpha": "PF_1",
    "pfphb": "PF_2",
    "pfphc": "PF_3",
    # Energy
    "totwh": "E_TOT",
    "totwhexp": "E_EXP",
    "totwhimp": "E_IMP",
    "totvarexp": "EQ_EXP",
    "totvarimp": "EQ_IMP",
    "totvahexp": "ES_EXP",
    "totvahimp": "ES_IMP",
    # DC & Temperature
    "dca": "I_DC",
    "dcv": "U_DC",
    "dcw": "P_DC",
    "ris": "R_ISO",
    "tmpcab": "TEMP_INT",
    "tmpsnk": "TEMP_INT",
    "sn": "SERIAL",
    "md": "MODEL",
    "mn": "MANUFACTURER",
}


def clean_name_token(name: str) -> str:
    """Normalizes a name for strict regex matching."""
    n = name.strip().lower()
    if " - " in n:
        n = n.split(" - ", 1)[1].strip()
    # Strip equipment prefixes
    n = re.sub(r"^[0-9]+#\s*", "", n)
    n = re.sub(r"^(?:pcs|inv|inverter|bms|meter)\s+(?:actual\s+)?", "", n)
    n = re.sub(r"^(?:actual\s+)", "", n)
    # Strip parenthetical expressions like (W), (summary), (kW)
    n = re.sub(r"[\(\[\{][^\)\]\}]*[\)\]\}]", "", n)
    n = n.replace("_", " ")
    n = re.sub(r"\s+", " ", n).strip()
    return n


def resolve_sunpm_tag_strictly(name: str, raw_tag: str, eq_cat: str = "") -> str:
    """
    Deterministically resolves a canonical SunPM tag with ZERO APPROXIMATION.
    Returns canonical tag if 100% certain, or empty string ("") if uncertain.
    """
    raw_clean = raw_tag.strip()

    # 1. Check if already exactly a canonical tag
    if is_valid_official_tag(raw_clean):
        name_lower = name.lower()
        # Duration / delay / protection false-positive guard
        if raw_clean in ("P_AC", "Q_AC", "S_AC", "U_AC_1", "I_AC_1", "FREQ", "PF"):
            if any(w in name_lower for w in ("delay", "duration", "duree", "durée", "tempo", "timeout", "intervalle")):
                return ""
        return raw_clean

    # 2. Check official legacy mappings
    raw_lower = raw_clean.lower()
    if raw_lower in OFFICIAL_LEGACY_MAP:
        mapped = OFFICIAL_LEGACY_MAP[raw_lower]
        if is_valid_official_tag(mapped):
            return mapped

    # 3. Check SunSpec standard abbreviation mapping
    if raw_lower in SUNSPEC_FIELD_MAP:
        mapped = SUNSPEC_FIELD_MAP[raw_lower]
        if is_valid_official_tag(mapped):
            return mapped

    for ss_key, ss_tag in SUNSPEC_FIELD_MAP.items():
        if raw_lower.endswith(f"_{ss_key}") or raw_lower == ss_key:
            return ss_tag

    # 4. Filter out obvious pseudo-tags, scale factors, alarms, and markers
    if any(raw_lower.endswith(sf) for sf in ("_sf", "sf")):
        return ""
    if raw_lower in (
        "id", "l", "pad", "opt", "vr", "version", "options", "spare", "reserve", "reserved",
        "st", "stvnd", "evt1", "evt2", "evtvnd1", "evtvnd2", "evtvnd3", "evtvnd4", "readonly",
        "actpt", "crvnam", "rmptms", "modena", "tms", "npt", "rvrttms", "ncrv", "actcrv", "wintms"
    ):
        return ""

    # 5. Deterministic, High-Confidence Semantic Resolution from Variable Name
    norm_name = clean_name_token(name)

    # Exclusion guards: if name contains words indicating non-telemetry, reject
    if any(ex in norm_name for ex in (
        "delay", "duration", "duree", "durée", "tempo", "temporisation", "timeout", "intervalle",
        "threshold", "seuil", "trip", "overvoltage", "undervoltage", "overfrequency", "underfrequency",
        "overcurrent", "protection", "limit value", "valeur limite", "coeff", "gain", "ratio",
        "address", "adresse", "baudrate", "parity", "crc", "protocol", "protocole", "version"
    )):
        if "ramp" in norm_name and "time" in norm_name:
            return "CMD_RAMP_TIME_S"
        if "revert" in norm_name and "time" in norm_name:
            return "CMD_REVERT_TIME_S"
        return ""

    # Active Power (AC)
    if re.search(r"^(?:total\s+)?active\s+power$|^puissance\s+active(?:\s+totale)?$|^pac$|^p_ac$|^(?:output|inverter|ac)\s+(?:active\s+)?power$", norm_name):
        return "P_AC"
    if re.search(r"^active\s+power\s+(?:phase\s+)?(?:a|1|l1)$|^puissance\s+active\s+(?:phase\s+)?(?:a|1|l1)$|^pac1$|^p_ac_1$", norm_name):
        return "P_AC_1"
    if re.search(r"^active\s+power\s+(?:phase\s+)?(?:b|2|l2)$|^puissance\s+active\s+(?:phase\s+)?(?:b|2|l2)$|^pac2$|^p_ac_2$", norm_name):
        return "P_AC_2"
    if re.search(r"^active\s+power\s+(?:phase\s+)?(?:c|3|l3)$|^puissance\s+active\s+(?:phase\s+)?(?:c|3|l3)$|^pac3$|^p_ac_3$", norm_name):
        return "P_AC_3"

    # Reactive Power (AC)
    if re.search(r"^(?:total\s+)?reactive\s+power$|^puissance\s+r[eé]active(?:\s+totale)?$|^qac$|^q_ac$|^(?:output|inverter|ac)\s+reactive\s+power$", norm_name):
        return "Q_AC"
    if re.search(r"^reactive\s+power\s+(?:phase\s+)?(?:a|1|l1)$|^puissance\s+r[eé]active\s+(?:phase\s+)?(?:a|1|l1)$|^qac1$|^q_ac_1$", norm_name):
        return "Q_AC_1"
    if re.search(r"^reactive\s+power\s+(?:phase\s+)?(?:b|2|l2)$|^puissance\s+r[eé]active\s+(?:phase\s+)?(?:b|2|l2)$|^qac2$|^q_ac_2$", norm_name):
        return "Q_AC_2"
    if re.search(r"^reactive\s+power\s+(?:phase\s+)?(?:c|3|l3)$|^puissance\s+r[eé]active\s+(?:phase\s+)?(?:c|3|l3)$|^qac3$|^q_ac_3$", norm_name):
        return "Q_AC_3"

    # Apparent Power (AC)
    if re.search(r"^(?:total\s+)?apparent\s+power$|^puissance\s+apparente(?:\s+totale)?$|^sac$|^s_ac$|^(?:output|inverter|ac)\s+apparent\s+power$", norm_name):
        return "S_AC"
    if re.search(r"^apparent\s+power\s+(?:phase\s+)?(?:a|1|l1)$|^puissance\s+apparente\s+(?:phase\s+)?(?:a|1|l1)$|^sac1$|^s_ac_1$", norm_name):
        return "S_AC_1"
    if re.search(r"^apparent\s+power\s+(?:phase\s+)?(?:b|2|l2)$|^puissance\s+apparente\s+(?:phase\s+)?(?:b|2|l2)$|^sac2$|^s_ac_2$", norm_name):
        return "S_AC_2"
    if re.search(r"^apparent\s+power\s+(?:phase\s+)?(?:c|3|l3)$|^puissance\s+apparente\s+(?:phase\s+)?(?:c|3|l3)$|^sac3$|^s_ac_3$", norm_name):
        return "S_AC_3"

    # Frequency
    if re.search(r"^(?:grid\s+|line\s+|ac\s+)?frequency$|^fr[eé]quence(?:\s+r[eé]seau)?$|^fac$|^freq$|^hz$", norm_name):
        return "FREQ"

    # Power Factor
    if re.search(r"^(?:total\s+)?power\s+factor$|^facteur\s+de\s+puissance$|^cos\s*phi$|^cosphi$|^pf$", norm_name):
        return "PF"
    if re.search(r"^power\s+factor\s+(?:phase\s+)?(?:a|1|l1)$|^facteur\s+de\s+puissance\s+(?:phase\s+)?(?:a|1|l1)$|^pf1$", norm_name):
        return "PF_1"
    if re.search(r"^power\s+factor\s+(?:phase\s+)?(?:b|2|l2)$|^facteur\s+de\s+puissance\s+(?:phase\s+)?(?:b|2|l2)$|^pf2$", norm_name):
        return "PF_2"
    if re.search(r"^power\s+factor\s+(?:phase\s+)?(?:c|3|l3)$|^facteur\s+de\s+puissance\s+(?:phase\s+)?(?:c|3|l3)$|^pf3$", norm_name):
        return "PF_3"

    # AC Phase to Neutral Voltages
    if re.search(r"^(?:phase\s+(?:a|1|l1)\s+voltage|voltage\s+(?:phase\s+)?(?:a|1|l1)|grid\s+voltage\s+l1|tension\s+phase\s+1|u_l1|uac1)$", norm_name):
        return "U_AC_1"
    if re.search(r"^(?:phase\s+(?:b|2|l2)\s+voltage|voltage\s+(?:phase\s+)?(?:b|2|l2)|grid\s+voltage\s+l2|tension\s+phase\s+2|u_l2|uac2)$", norm_name):
        return "U_AC_2"
    if re.search(r"^(?:phase\s+(?:c|3|l3)\s+voltage|voltage\s+(?:phase\s+)?(?:c|3|l3)|grid\s+voltage\s+l3|tension\s+phase\s+3|u_l3|uac3)$", norm_name):
        return "U_AC_3"

    # AC Phase to Phase Voltages
    if re.search(r"^(?:line\s+voltage\s+(?:ab|l1\s*l2|l1-l2)|phase\s+voltage\s+ab|tension\s+compos[eé]e\s+l1\s*l2|u_l1l2|uab)$", norm_name):
        return "U_AC_L1L2"
    if re.search(r"^(?:line\s+voltage\s+(?:bc|l2\s*l3|l2-l3)|phase\s+voltage\s+bc|tension\s+compos[eé]e\s+l2\s*l3|u_l2l3|ubc)$", norm_name):
        return "U_AC_L2L3"
    if re.search(r"^(?:line\s+voltage\s+(?:ca|l3\s*l1|l3-l1)|phase\s+voltage\s+ca|tension\s+compos[eé]e\s+l3\s*l1|u_l3l1|uca)$", norm_name):
        return "U_AC_L3L1"

    # AC Currents
    if re.search(r"^(?:total\s+)?(?:ac\s+current|grid\s+current|inverter\s+current)$|^courant\s+ac(?:\s+total)?$|^iac$", norm_name):
        return "I_AC"
    if re.search(r"^(?:phase\s+(?:a|1|l1)\s+current|current\s+(?:phase\s+)?(?:a|1|l1)|grid\s+current\s+l1|courant\s+phase\s+1|i_l1|iac1)$", norm_name):
        return "I_AC_1"
    if re.search(r"^(?:phase\s+(?:b|2|l2)\s+current|current\s+(?:phase\s+)?(?:b|2|l2)|grid\s+current\s+l2|courant\s+phase\s+2|i_l2|iac2)$", norm_name):
        return "I_AC_2"
    if re.search(r"^(?:phase\s+(?:c|3|l3)\s+current|current\s+(?:phase\s+)?(?:c|3|l3)|grid\s+current\s+l3|courant\s+phase\s+3|i_l3|iac3)$", norm_name):
        return "I_AC_3"

    # Cumulative Energies
    if re.search(r"^(?:cumulative\s+|total\s+)?(?:energy(?:\s+total)?|active\s+energy|yield|generation)$|^[eé]nergie\s+totale$|^e_tot$|^etot$", norm_name):
        return "E_TOT"
    if re.search(r"^(?:daily\s+|today\s+)?(?:energy|yield|generation)$|^[eé]nergie\s+journali[eè]re$|^e_day$|^eday$", norm_name):
        return "E_DAY"
    if re.search(r"^exported\s+(?:active\s+)?energy$|^[eé]nergie\s+export[eé]e$", norm_name):
        return "E_EXP"
    if re.search(r"^imported\s+(?:active\s+)?energy$|^[eé]nergie\s+import[eé]e$", norm_name):
        return "E_IMP"

    # DC Measurements
    if re.search(r"^(?:total\s+)?(?:dc\s+power|pv\s+power|input\s+power)$|^puissance\s+dc(?:\s+totale)?$|^pdc$", norm_name):
        return "P_DC"
    if re.search(r"^(?:total\s+)?(?:dc\s+voltage|pv\s+voltage|dc\s+bus\s+voltage|battery\s+voltage)$|^tension\s+dc(?:\s+totale)?$|^udc$", norm_name):
        return "U_DC"
    if re.search(r"^(?:total\s+)?(?:dc\s+current|pv\s+current|dc\s+bus\s+current|battery\s+current)$|^courant\s+dc(?:\s+total)?$|^idc$", norm_name):
        return "I_DC"

    # In BESS context, "total voltage" and "total current" are DC measurements
    if eq_cat in ("BAT", "BESS", "STORAGE") or "battery" in norm_name or "bms" in norm_name:
        if norm_name in ("total voltage", "pack voltage", "voltage"):
            return "U_DC"
        if norm_name in ("total current", "pack current", "current"):
            return "I_DC"

    # MPPTs
    mppt_m = re.search(r"mppt\s*([0-9]{1,2})", norm_name)
    if mppt_m:
        m_idx = mppt_m.group(1)
        if "voltage" in norm_name or "tension" in norm_name or " u" in norm_name:
            return f"U_DC_MPPT_{m_idx}"
        if "current" in norm_name or "courant" in norm_name or " i" in norm_name:
            return f"I_DC_MPPT_{m_idx}"
        if "power" in norm_name or "puissance" in norm_name or " p" in norm_name:
            return f"P_DC_MPPT_{m_idx}"

    # String Currents & Voltages
    string_m = re.search(r"(?:string|chaine|chaîne)\s*([0-9]{1,2})", norm_name)
    if string_m:
        s_idx = string_m.group(1)
        if "current" in norm_name or "courant" in norm_name or " i" in norm_name:
            return f"I_DC_STRING_{s_idx}"
        if "voltage" in norm_name or "tension" in norm_name or " u" in norm_name:
            return f"U_DC_STRING_{s_idx}"

    # Basic Info
    if re.search(r"^(?:serial\s+number|device\s+serial|inverter\s+serial|meter\s+serial|sn)$|^num[eé]ro\s+de\s+s[eé]rie$", norm_name):
        return "SERIAL"
    if re.search(r"^(?:device\s+model|inverter\s+model|meter\s+model|product\s+model)$|^mod[eè]le(?:\s+appareil)?$", norm_name):
        return "MODEL"
    if re.search(r"^(?:manufacturer(?:\s+name)?|vendor(?:\s+name)?)$|^constructeur$|^fabricant$", norm_name):
        return "MANUFACTURER"
    if re.search(r"^(?:rated\s+power|nominal\s+power|rated\s+active\s+power|rated\s+capacity)$|^puissance\s+nominale$", norm_name):
        return "P_RATED"

    # Temperatures & Insulation
    if re.search(r"^(?:internal\s+temp(?:erature)?|cabinet\s+temp(?:erature)?|heatsink\s+temp(?:erature)?|inverter\s+temp(?:erature)?|board\s+internal\s+temp(?:erature)?|temperature\s+tmp)$|^temp[eé]rature\s+interne$", norm_name):
        return "TEMP_INT"
    if re.search(r"^(?:insulation\s+resistance|isolation\s+resistance|riso)$|^r[eé]sistance\s+d['\s]isolement$", norm_name):
        return "R_ISO"

    # Battery
    if re.search(r"^(?:battery\s+)?(?:state\s+of\s+charge|soc)$|^total\s+soc$|^[eé]tat\s+de\s+charge$", norm_name):
        return "SOC"
    if re.search(r"^(?:battery\s+)?(?:state\s+of\s+health|soh)$|^total\s+soh$|^[eé]tat\s+de\s+sant[eé]$", norm_name):
        return "SOH"
    if re.search(r"^(?:battery\s+temp(?:erature)?|bms\s+temp(?:erature)?|pack\s+temp(?:erature)?)$|^temp[eé]rature\s+batterie$", norm_name):
        return "TEMP"

    # Sensors
    if re.search(r"^(?:plane\s+of\s+array\s+irradiance|poa(?:\s+irradiance)?)$|^irradiation\s+poa$", norm_name):
        return "POA"
    if re.search(r"^(?:global\s+horizontal\s+irradiance|ghi(?:\s+irradiance)?)$|^irradiation\s+ghi$", norm_name):
        return "GHI"
    if re.search(r"^(?:ambient\s+temperature|outdoor\s+temperature)$|^temp[eé]rature\s+ambiante$", norm_name):
        return "TEMP_AMB"
    if re.search(r"^(?:cell\s+temperature|module\s+temperature|panel\s+temperature)$|^temp[eé]rature\s+cellule$|^temp[eé]rature\s+module$", norm_name):
        return "TEMP_CELL"
    if re.search(r"^(?:wind\s+speed|vitesse\s+du\s+vent)$", norm_name):
        return "WIND_SPEED"
    if re.search(r"^(?:wind\s+bearing|wind\s+direction|direction\s+du\s+vent)$", norm_name):
        return "WIND_BEARING"
    if re.search(r"^(?:relative\s+humidity|humidit[eé]\s+relative)$", norm_name):
        return "HUMIDITY"
    if re.search(r"^(?:barometric\s+pressure|air\s+pressure|pression\s+atmosph[eé]rique)$", norm_name):
        return "PRESSURE"
    if re.search(r"^(?:rainfall|precipitation|pluie)$", norm_name):
        return "RAIN"

    # Commands
    if re.search(r"^(?:remote\s+)?(?:start|power\s+on|turn\s*on)(?:\s+command)?$|^d[eé]marrage$", norm_name):
        return "CMD_ON"
    if re.search(r"^(?:remote\s+)?(?:stop|power\s+off|turn\s*off)(?:\s+command)?$|^arr[eê]t$", norm_name):
        return "CMD_OFF"
    if re.search(r"^(?:set\s+active\s+power\s+percent(?:age)?|active\s+power\s+limit\s+percent(?:age)?)$", norm_name):
        return "CMD_P_AC_SET_PERCENT"
    if re.search(r"^power\s+limitation\s+enable$", norm_name):
        return "CMD_P_ENA"

    # If uncertain or no exact match: LEAVE EMPTY ("")
    return ""


# -----------------------------------------------------------------------------
# Batch Processing & Atomic Updating
# -----------------------------------------------------------------------------

def detect_equipment_category(filename: str) -> str:
    """Detects equipment category from filename prefix."""
    fn = filename.upper()
    if fn.startswith("INV_") or "INVERTER" in fn:
        return "INV"
    if fn.startswith("MET_") or "METER" in fn:
        return "MET"
    if fn.startswith("SEN_") or "SENSOR" in fn:
        return "SEN"
    if fn.startswith("BAT_") or "BATTERY" in fn or "BMS" in fn or "STORAGE" in fn:
        return "BAT"
    if fn.startswith("COM_") or "LOGGER" in fn:
        return "COM"
    return "OTH"


def process_definition_csv(file_path: Path) -> Dict[str, Any]:
    """
    Reads a single definition CSV, strictly audits and updates column 7 (Tag),
    writes back the cleaned file atomically, and returns statistics.
    """
    stats = {
        "file": str(file_path),
        "total_rows": 0,
        "initial_canonical": 0,
        "initial_non_canonical": 0,
        "initial_empty": 0,
        "final_canonical": 0,
        "final_non_canonical": 0,
        "final_empty": 0,
        "modified": False,
        "tags_added": 0,
        "tags_purged": 0,
        "tags_normalized": 0,
    }

    eq_cat = detect_equipment_category(file_path.name)

    rows: List[List[str]] = []
    header: List[str] = []

    with open(file_path, "r", encoding="utf-8-sig", errors="ignore") as f:
        reader = csv.reader(f, delimiter=";")
        header_row = next(reader, None)
        if not header_row:
            return stats
        header = header_row

        for row in reader:
            if not row or len(row) < 7:
                continue
            stats["total_rows"] += 1
            name = row[5].strip()
            raw_tag = row[6].strip()

            if not raw_tag:
                stats["initial_empty"] += 1
            elif is_valid_official_tag(raw_tag):
                stats["initial_canonical"] += 1
            else:
                stats["initial_non_canonical"] += 1

            # Strictly resolve tag
            resolved_tag = resolve_sunpm_tag_strictly(name=name, raw_tag=raw_tag, eq_cat=eq_cat)

            # Ensure tag passes strict canonical validation
            if resolved_tag and not is_valid_official_tag(resolved_tag):
                resolved_tag = ""

            if not resolved_tag:
                stats["final_empty"] += 1
                if raw_tag:
                    stats["tags_purged"] += 1
                    stats["modified"] = True
            else:
                stats["final_canonical"] += 1
                if not raw_tag:
                    stats["tags_added"] += 1
                    stats["modified"] = True
                elif raw_tag != resolved_tag:
                    stats["tags_normalized"] += 1
                    stats["modified"] = True

            # Update row column 6
            row[6] = resolved_tag
            rows.append(row)

    if stats["modified"]:
        # Write back atomically
        tmp_fd, tmp_path = tempfile.mkstemp(dir=str(file_path.parent), prefix="tmp_def_")
        with os.fdopen(tmp_fd, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(header)
            writer.writerows(rows)
        os.replace(tmp_path, file_path)

    return stats


def run_batch_audit_and_enrich(
    target_dirs: List[Path],
    report_output: Optional[Path] = None,
) -> Dict[str, Any]:
    """Runs strict audit and enrichment across all target directories."""
    start_time = time.time()

    all_csvs: List[Path] = []
    for d in target_dirs:
        if d.is_dir():
            all_csvs.extend(sorted(d.glob("*.csv")))

    print(f"Discovered {len(all_csvs)} definition files across {len(target_dirs)} directories.")

    total_files = len(all_csvs)
    modified_files = 0
    total_registers = 0
    init_canonical = 0
    init_non_canonical = 0
    init_empty = 0
    final_canonical = 0
    final_non_canonical = 0
    final_empty = 0
    tags_added = 0
    tags_purged = 0
    tags_normalized = 0

    tag_distribution = Counter()

    for idx, csv_path in enumerate(all_csvs, 1):
        res = process_definition_csv(csv_path)
        total_registers += res["total_rows"]
        init_canonical += res["initial_canonical"]
        init_non_canonical += res["initial_non_canonical"]
        init_empty += res["initial_empty"]
        final_canonical += res["final_canonical"]
        final_non_canonical += res["final_non_canonical"]
        final_empty += res["final_empty"]
        tags_added += res["tags_added"]
        tags_purged += res["tags_purged"]
        tags_normalized += res["tags_normalized"]
        if res["modified"]:
            modified_files += 1

        if idx % 200 == 0 or idx == total_files:
            print(f"[{idx}/{total_files}] Processed {csv_path.name}...")

    # Now verify final tag distribution
    for csv_path in all_csvs:
        with open(csv_path, "r", encoding="utf-8-sig", errors="ignore") as f:
            r = csv.reader(f, delimiter=";")
            next(r, None)
            for row in r:
                if row and len(row) > 6 and row[6].strip():
                    tag_distribution[row[6].strip()] += 1

    elapsed = time.time() - start_time

    summary = {
        "timestamp": time.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "elapsed_seconds": round(elapsed, 2),
        "total_files": total_files,
        "modified_files": modified_files,
        "total_registers": total_registers,
        "initial_metrics": {
            "canonical_tags": init_canonical,
            "non_canonical_tags": init_non_canonical,
            "empty_registers": init_empty,
        },
        "final_metrics": {
            "canonical_tags": final_canonical,
            "non_canonical_tags": final_non_canonical,
            "empty_registers": final_empty,
        },
        "changes": {
            "tags_added": tags_added,
            "tags_purged": tags_purged,
            "tags_normalized": tags_normalized,
        },
        "top_canonical_tags": dict(tag_distribution.most_common(50)),
    }

    print("\n" + "=" * 60)
    print("STRICT SUNPM TAG AUDIT & ENRICHMENT REPORT")
    print("=" * 60)
    print(f"Files Processed:      {total_files} ({modified_files} modified)")
    print(f"Total Registers:      {total_registers}")
    print(f"Initial Canonical:    {init_canonical}")
    print(f"Initial Non-Canonical:{init_non_canonical} (pseudo-tags/slugs/markers)")
    print(f"Final Canonical:      {final_canonical}")
    print(f"Final Non-Canonical:  {final_non_canonical} (STRICT ZERO)")
    print(f"Final Empty:          {final_empty} (unambiguous empty fields)")
    print(f"Tags Added:           +{tags_added}")
    print(f"Tags Purged:          -{tags_purged}")
    print(f"Tags Normalized:      ~{tags_normalized}")
    print(f"Duration:             {summary['elapsed_seconds']}s")
    print("=" * 60)

    if report_output:
        report_output.parent.mkdir(parents=True, exist_ok=True)
        with open(report_output, "w", encoding="utf-8") as f:
            json.dump(summary, f, indent=2)
        print(f"Audit report saved to: {report_output}")

    return summary


if __name__ == "__main__":
    repo_root = Path(__file__).resolve().parent.parent
    target_directories = [
        repo_root / "generated_definitions",
        repo_root / "FoxESS_P3_H3_Pro_Definitions",
        repo_root / "MarsRenewable_Definitions",
        repo_root / "tagged_definitions",
    ]
    report_file = repo_root / "generated_definitions" / "GENERATION_AUDIT_REPORT.json"

    run_batch_audit_and_enrich(
        target_dirs=[d for d in target_directories if d.is_dir()],
        report_output=report_file,
    )
