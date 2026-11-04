#!/usr/bin/env python3
"""
Upgrade Huawei definitions to the complete superset (392 registers).
Reads the authoritative 392-register definition from Definitions/INV_Huawei_V4.csv
and distributes it to all Huawei V4 variants in generated_definitions/.
"""

import csv
import sys
from pathlib import Path

WORKSPACE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(WORKSPACE_DIR))
sys.path.insert(0, str(WORKSPACE_DIR / "scratch"))

from DefFileGenerator.def_gen import Generator

v4_ref_path = WORKSPACE_DIR / "Definitions" / "INV_Huawei_V4.csv"

with open(v4_ref_path, "r", encoding="utf-8-sig", errors="ignore") as f:
    reader = csv.reader(f, delimiter=";")
    header = next(reader)
    final_rows = [r for r in reader if len(r) > 6]

print(f"Authoritative Huawei V4 registers: {len(final_rows)}")

target_files = [
    WORKSPACE_DIR / "generated_definitions" / "INV_Huawei_V4.csv",
    WORKSPACE_DIR / "generated_definitions" / "INV_Huawei_SUN2000_V4.csv",
    WORKSPACE_DIR / "generated_definitions" / "INV_Huawei_SUN2000MA_V100R001C00.csv",
    WORKSPACE_DIR / "generated_definitions" / "INV_Webdyn_HUAWEI_V4.csv",
    WORKSPACE_DIR / "generated_definitions" / "INV_Webdyn_HUAWEI_ModbusV4.csv",
    WORKSPACE_DIR / "generated_definitions" / "INV_Webdyn_WPM015004_HUAWEI_V4.csv",
    WORKSPACE_DIR / "generated_definitions" / "INV_Webdyn_WPMXXXXX_modbusRTU_Inverter_HUAWEI_V4.csv",
    WORKSPACE_DIR / "generated_definitions" / "OTH_Huawei_HUAWEI_V4.csv",
    WORKSPACE_DIR / "generated_definitions" / "OTH_Webdyn_HUAWEI_V4.csv",
    WORKSPACE_DIR / "generated_definitions" / "OTH_Webdyn_HUAWEI_ModbusV4.csv",
    WORKSPACE_DIR / "generated_definitions" / "OTH_Webdyn_WPM015004_HUAWEI_V4.csv",
]

gen = Generator()

for tf in target_files:
    tf.parent.mkdir(parents=True, exist_ok=True)
    file_header = list(header)
    if "SUN2000MA" in tf.name:
        file_header[3] = "SUN2000MA_V100R001C00"
    elif "WPM" in tf.name:
        file_header[3] = tf.stem
    else:
        file_header[3] = "SUN2000_V4"

    with open(tf, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(file_header)
        writer.writerows(final_rows)

    rep = gen.validate_csv_detailed(str(tf), strict=False, strict_overlap=False)
    print(f"Updated {tf.name}: {len(final_rows)} regs, valid={rep.is_valid}")

# Upgrade COM_Webdyn_HUAWEI_SmartLogger3000.csv if present
full_sl3000 = WORKSPACE_DIR / "generated_definitions" / "COM_Huawei_SmartLogger3000.csv"
short_sl3000 = WORKSPACE_DIR / "generated_definitions" / "COM_Webdyn_HUAWEI_SmartLogger3000.csv"
if full_sl3000.is_file() and short_sl3000.is_file():
    with open(full_sl3000, "r", encoding="utf-8", errors="ignore") as f:
        sl_lines = f.readlines()
    with open(short_sl3000, "w", encoding="utf-8", newline="") as f:
        f.writelines(sl_lines)
    print(f"Updated COM_Webdyn_HUAWEI_SmartLogger3000.csv to {len(sl_lines)-1} regs")

# Upgrade INV_Huawei_81012141820.csv if test source was present, then purge test file
test_810 = WORKSPACE_DIR / "generated_definitions" / "INV_Huawei_81012141820_Test.csv"
short_810 = WORKSPACE_DIR / "generated_definitions" / "INV_Huawei_81012141820.csv"
if test_810.is_file():
    if short_810.is_file():
        with open(test_810, "r", encoding="utf-8", errors="ignore") as f:
            lines_810 = f.readlines()
        with open(short_810, "w", encoding="utf-8", newline="") as f:
            f.writelines(lines_810)
        print(f"Updated INV_Huawei_81012141820.csv to {len(lines_810)-1} regs")
    # Delete test file so only clean production definitions remain
    test_810.unlink(missing_ok=True)

print("All Huawei definitions successfully synchronized to full superset!")
