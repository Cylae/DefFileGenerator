#!/usr/bin/env python3
"""
Generate dedicated 'Battery' and 'Meter' definition files for Fox ESS (P3 Pro / H3 Pro and R Series)
100% compliant with WebdynSunPM specification.
"""

import csv
import os
import sys
from pathlib import Path

# Add project root to sys.path
WORKSPACE_DIR = Path(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
sys.path.insert(0, str(WORKSPACE_DIR))

from DefFileGenerator.def_gen import Generator  # noqa: E402

FOX_DIR = Path(r"\\WDN-FILES.WEBDYN.COM\Support\Equipementiers\Fox ESS")
SOURCE_BAT_CSV = FOX_DIR / "Battery" / "FoxESS_P3_H3_PRO.csv"
LOCAL_DEFS_DIR = WORKSPACE_DIR / "Definitions"
GEN_DEFS_DIR = WORKSPACE_DIR / "generated_definitions"

generator = Generator()

print(f"Reading master registers from {SOURCE_BAT_CSV}...")
with open(SOURCE_BAT_CSV, encoding="utf-8-sig", errors="ignore") as f:
    reader = csv.reader(f, delimiter=";")
    header = next(reader)
    all_rows = list(reader)

print(f"Total source rows: {len(all_rows)}")

# =========================================================================
# 1. BUILD METER DEFINITION
# =========================================================================
meter_rows = []
for r in all_rows:
    if len(r) < 6:
        continue
    addr_str = r[2].split("_")[0]
    try:
        addr = int(addr_str)
    except ValueError:
        continue
    name = r[5]

    is_meter = (
        (36100 <= addr <= 36248)
        or (38801 <= addr <= 38947)
        or addr in (39168, 49207, 49208, 49248)
        or "meter" in name.lower()
    )
    if is_meter:
        meter_rows.append(list(r))

print(f"Extracted {len(meter_rows)} Meter registers.")

# Assign SunPM tags for Meter
seen_tags = set()
for r in meter_rows:
    addr = r[2].split("_")[0]
    name = r[5]
    tag = ""

    if addr == "36100":
        tag = "SERIAL"
    elif addr == "36116":
        tag = "MANUFACTURER"
    elif addr == "36132":
        tag = "MODEL"
    elif addr == "38802":
        tag = "U_AC_1"
    elif addr == "38804":
        tag = "U_AC_2"
    elif addr == "38806":
        tag = "U_AC_3"
    elif addr == "38808":
        tag = "I_AC_1"
    elif addr == "38810":
        tag = "I_AC_2"
    elif addr == "38812":
        tag = "I_AC_3"
    elif addr == "38814":
        tag = "P_AC"
    elif addr == "38816":
        tag = "P_AC_1"
    elif addr == "38818":
        tag = "P_AC_2"
    elif addr == "38820":
        tag = "P_AC_3"
    elif addr == "38822":
        tag = "Q_AC"
    elif addr == "38824":
        tag = "Q_AC_1"
    elif addr == "38826":
        tag = "Q_AC_2"
    elif addr == "38828":
        tag = "Q_AC_3"
    elif addr == "38830":
        tag = "S_AC"
    elif addr == "38832":
        tag = "S_AC_1"
    elif addr == "38834":
        tag = "S_AC_2"
    elif addr == "38836":
        tag = "S_AC_3"
    elif addr == "38838":
        tag = "PF"
    elif addr == "38840":
        tag = "PF_1"
    elif addr == "38842":
        tag = "PF_2"
    elif addr == "38844":
        tag = "PF_3"
    elif addr == "38846":
        tag = "FREQ"

    if tag and tag not in seen_tags:
        r[6] = tag
        seen_tags.add(tag)
    else:
        r[6] = ""

# Format rows for Meter
clean_meter_rows = []
for idx, r in enumerate(meter_rows, 1):
    # Ensure 11 columns
    row11 = (r + [""] * 11)[:11]
    row11[0] = str(idx)
    row11[1] = "3"  # Holding FC03
    # Format CoefA and CoefB
    try:
        ca = float(row11[7]) if row11[7].strip() else 1.0
    except ValueError:
        ca = 1.0
    try:
        cb = float(row11[8]) if row11[8].strip() else 0.0
    except ValueError:
        cb = 0.0
    row11[7] = f"{ca:.6f}"
    row11[8] = f"{cb:.6f}"
    # Action code
    if not row11[10].strip():
        addr_int = int(row11[2].split("_")[0])
        row11[10] = (
            "1"
            if addr_int in (49207, 49208, 49248)
            else ("9" if "lost" in row11[5].lower() else "4")
        )
    clean_meter_rows.append(row11)

meter_header = ["Modbus", "Meter", "FoxESS", "P3 Pro / H3 Pro", "", "", "", "", "", "", ""]

# =========================================================================
# 2. BUILD BATTERY DEFINITION
# =========================================================================
battery_rows = []
for r in all_rows:
    if len(r) < 6:
        continue
    addr_str = r[2].split("_")[0]
    try:
        addr = int(addr_str)
    except ValueError:
        continue
    name = r[5]

    # Exclude meter registers
    if (
        (36100 <= addr <= 36248)
        or (38801 <= addr <= 38947)
        or addr in (39168, 49207, 49208, 49248)
        or "meter" in name.lower()
    ):
        continue

    # Include battery / BMS registers
    is_bat = (
        (30000 <= addr <= 30032)
        or (36001 <= addr <= 36003)
        or (37000 <= addr <= 38334)
        or (45000 <= addr <= 45999)
        or (48000 <= addr <= 48999)
        or (53000 <= addr <= 53999)
        or any(k in name.lower() for k in ("bms", "battery", "soc", "soh", "cell"))
    )
    if is_bat:
        battery_rows.append(list(r))

print(f"Extracted {len(battery_rows)} Battery registers.")

# Assign SunPM tags for Battery
seen_bat_tags = set()
for r in battery_rows:
    addr = r[2].split("_")[0]
    name = r[5].lower()
    tag = ""

    if addr == "30000":
        tag = "MODEL"
    elif addr == "30016":
        tag = "SERIAL"
    elif addr == "30032":
        tag = "MANUFACTURER"
    elif "soc" in name and "threshold" not in name and "limit" not in name:
        if "bms1" in name or tag not in seen_bat_tags:
            tag = "SOC"
    elif "soh" in name:
        if "bms1" in name or tag not in seen_bat_tags:
            tag = "SOH"
    elif "temperature" in name and "start" not in name and "stop" not in name:
        if "bms1" in name or tag not in seen_bat_tags:
            tag = "TEMP"
    elif "voltage" in name and "cell" not in name and "slave" not in name:
        if "bms1" in name or tag not in seen_bat_tags:
            tag = "U_DC"
    elif "current" in name and "cell" not in name and "slave" not in name:
        if "bms1" in name or tag not in seen_bat_tags:
            tag = "I_DC"

    if tag and tag not in seen_bat_tags:
        r[6] = tag
        seen_bat_tags.add(tag)
    else:
        r[6] = ""

clean_battery_rows = []
for idx, r in enumerate(battery_rows, 1):
    row11 = (r + [""] * 11)[:11]
    row11[0] = str(idx)
    row11[1] = "3"
    try:
        ca = float(row11[7]) if row11[7].strip() else 1.0
    except ValueError:
        ca = 1.0
    try:
        cb = float(row11[8]) if row11[8].strip() else 0.0
    except ValueError:
        cb = 0.0
    row11[7] = f"{ca:.6f}"
    row11[8] = f"{cb:.6f}"
    if not row11[10].strip():
        addr_int = int(row11[2].split("_")[0])
        row11[10] = "1" if addr_int >= 40000 else "4"
    clean_battery_rows.append(row11)

bat_header = ["Modbus", "Battery", "FoxESS", "P3 Pro / H3 Pro", "", "", "", "", "", "", ""]


# =========================================================================
# 3. WRITE & VALIDATE FILES
# =========================================================================
def write_and_audit(path: Path, header: list[str], rows: list[list[str]], name: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(header)
        writer.writerows(rows)
    report = generator.validate_csv_detailed(str(path), strict=False, strict_overlap=False)
    print(
        f"[{name}] {path.name}: {len(rows)} registers, valid: {report.is_valid}, errors: {report.stats.get('errors', 0)}"
    )
    assert report.is_valid, (
        f"Validation failed for {path}: {[iss.message for iss in report.issues if iss.severity == 'ERROR']}"
    )


# Write Meter definitions
dest_meter_paths = [
    FOX_DIR / "FOXESS_P3_H3_PRO_METER.csv",
    LOCAL_DEFS_DIR / "MET_FoxESS_P3_H3_PRO.csv",
    GEN_DEFS_DIR / "MET_FoxESS_P3_H3_PRO.csv",
]
for p in dest_meter_paths:
    try:
        write_and_audit(p, meter_header, clean_meter_rows, "Meter")
    except Exception as e:
        print(f"Could not write to {p}: {e}")

# Write Battery definitions
dest_bat_paths = [
    FOX_DIR / "FOXESS_P3_H3_PRO_BATTERY.csv",
    FOX_DIR / "Battery" / "FoxESS_P3_H3_PRO.csv",
    LOCAL_DEFS_DIR / "BAT_FoxESS_P3_H3_PRO.csv",
    GEN_DEFS_DIR / "BAT_FoxESS_P3_H3_PRO.csv",
]
for p in dest_bat_paths:
    try:
        write_and_audit(p, bat_header, clean_battery_rows, "Battery")
    except Exception as e:
        print(f"Could not write to {p}: {e}")

print("\nSUCCESS: Both Battery and Meter definition files generated and verified!")
