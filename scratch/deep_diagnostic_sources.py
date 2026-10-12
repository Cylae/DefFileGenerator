#!/usr/bin/env python3
"""
Deep Diagnostic and Bug Hunting Runner across Real Equipment Manufacturer Sources.
Extracts tables, normalizes fields, checks for anomalies, warnings, and unhandled errors.
"""

import logging
import os
import sys
import traceback

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor


# Intercept and record logging warnings/errors
class WarningCollector(logging.Handler):
    def __init__(self):
        super().__init__()
        self.records = []

    def emit(self, record):
        if record.levelno >= logging.WARNING:
            self.records.append(
                {
                    "level": record.levelname,
                    "msg": record.getMessage(),
                    "logger": record.name,
                }
            )


collector = WarningCollector()
logging.getLogger().addHandler(collector)
logging.getLogger().setLevel(logging.WARNING)

EQUIP_DIR = r"G:\My Drive\08092026\Equipementiers"
MARS_SOURCE = r"C:\Users\Cylae\Downloads\MarsLocalController_IO List_ModbusTCP_V7.0.xlsx"

# Select candidate source files: All Excel spreadsheets + key Modbus PDFs
source_files = []

# 1. Mars
if os.path.exists(MARS_SOURCE):
    source_files.append(("Mars Energy", "MarsLocalController_V7.0.xlsx", MARS_SOURCE))

# 2. Excel spreadsheets in Equipementiers (Sources/ and subfolders)
for root, _dirs, files in os.walk(EQUIP_DIR):
    for f in files:
        ext = os.path.splitext(f)[1].lower()
        if ext in (".xlsx", ".xlsm") and not f.startswith("~$"):
            full_path = os.path.join(root, f)
            rel_path = os.path.relpath(full_path, EQUIP_DIR)
            mfg = rel_path.split(os.sep)[0]
            source_files.append((mfg, rel_path, full_path))

# 3. Dedicated Modbus protocol PDFs in Sources/
for root, _dirs, files in os.walk(EQUIP_DIR):
    for f in files:
        ext = os.path.splitext(f)[1].lower()
        if ext == ".pdf":
            f_low = f.lower()
            if any(
                k in f_low
                for k in (
                    "modbus",
                    "communication protocol",
                    "protocol modbus",
                    "interface definitions",
                    "io list",
                )
            ):
                if not any(
                    skip in f_low
                    for skip in ("declaration", "certificat", "brochure", "flyer", "catalog")
                ):
                    full_path = os.path.join(root, f)
                    rel_path = os.path.relpath(full_path, EQUIP_DIR)
                    mfg = rel_path.split(os.sep)[0]
                    source_files.append((mfg, rel_path, full_path))

print(f"Total source candidate files to test: {len(source_files)}")

results = []
ext = Extractor()
gen = Generator()

for idx, (mfg, rel, path) in enumerate(source_files, 1):
    collector.records.clear()
    res = {
        "index": idx,
        "mfg": mfg,
        "file": rel,
        "path": path,
        "status": "UNKNOWN",
        "registers_extracted": 0,
        "warnings": [],
        "errors": [],
        "exception": None,
    }

    try:
        file_ext = os.path.splitext(path)[1].lower()
        if file_ext in (".xlsx", ".xlsm"):
            raw_tables = list(ext.extract_from_excel(path))
        elif file_ext == ".pdf":
            raw_tables = list(ext.extract_from_pdf(path))
        else:
            continue

        cleaned = list(ext.map_and_clean(raw_tables))
        res["registers_extracted"] = len(cleaned)

        if not cleaned:
            res["status"] = "ZERO_REGISTERS"
        else:
            # Process with Generator to audit syntax, overlap, actions, and functions
            processed = list(gen.process_rows(cleaned))
            res["registers_processed"] = len(processed)
            res["status"] = "SUCCESS"

    except Exception as e:
        res["status"] = "EXCEPTION"
        res["exception"] = f"{type(e).__name__}: {str(e)}"
        res["traceback"] = traceback.format_exc()

    res["warnings"] = list(collector.records)
    results.append(res)

    status_icon = "PASS" if res["status"] == "SUCCESS" else res["status"]
    warn_count = len(res["warnings"])
    print(
        f"[{idx:3d}/{len(source_files):3d}] [{status_icon:14s}] {mfg:15s} | {os.path.basename(path):35s} | Regs: {res['registers_extracted']:5d} | Warn: {warn_count:2d}"
    )

# Summary
successes = [r for r in results if r["status"] == "SUCCESS"]
zeros = [r for r in results if r["status"] == "ZERO_REGISTERS"]
exceptions = [r for r in results if r["status"] == "EXCEPTION"]

print("\n" + "=" * 80)
print("DEEP DIAGNOSTIC AUDIT SUMMARY:")
print(f"Total documents tested : {len(results)}")
print(f"Successful extractions : {len(successes)}")
print(f"Zero registers detected: {len(zeros)}")
print(f"Exceptions / Crashes   : {len(exceptions)}")
print("=" * 80)

if exceptions:
    print("\nEXCEPTIONS DETAILS:")
    for ex in exceptions:
        print(f"File: {ex['file']}")
        print(f"Exception: {ex['exception']}")
        print(f"Traceback:\n{ex.get('traceback', '')[:500]}\n")

if zeros:
    print("\nZERO REGISTERS SAMPLES (First 15):")
    for z in zeros[:15]:
        print(f"  {z['mfg']} | {z['file']}")

# Warnings Analysis
all_warns = {}
for r in results:
    for w in r["warnings"]:
        msg = w["msg"]
        # Group by category prefix
        prefix = msg.split(":")[0] if ":" in msg else msg[:40]
        all_warns[prefix] = all_warns.get(prefix, 0) + 1

print("\nWARNING PATTERNS OBSERVED:")
for p, count in sorted(all_warns.items(), key=lambda x: -x[1])[:20]:
    print(f"  [{count:4d}x] {p}")
