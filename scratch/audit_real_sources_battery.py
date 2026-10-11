#!/usr/bin/env python3
import glob
import os
import sys
import traceback
from DefFileGenerator.extractor import Extractor

DATASET_ROOT = r"G:\My Drive\08092026\Equipementiers"


def audit_all():
    ext = Extractor()
    files = []
    for ext_pattern in ["*.xlsx", "*.xls", "*.csv"]:
        files.extend(glob.glob(os.path.join(DATASET_ROOT, "**", ext_pattern), recursive=True))

    print(f"Total candidate source files found: {len(files)}")

    results = {"success": [], "zero_registers": [], "error": []}

    for fpath in files:
        f_lower = fpath.lower()
        # skip non-technical non-modbus files
        if any(
            ign in f_lower
            for ign in [
                "ticket",
                "support",
                "reunion",
                "prio",
                "qualité",
                "client",
                "trame",
                "synthese",
            ]
        ):
            continue

        rel = os.path.relpath(fpath, DATASET_ROOT)
        ext_type = os.path.splitext(fpath)[1].lower()

        try:
            total_regs = 0
            tables_count = 0
            if ext_type in [".xlsx", ".xls"]:
                tables_gen = ext.extract_from_excel(fpath)
            elif ext_type == ".pdf":
                tables_gen = ext.extract_from_pdf(fpath)
            elif ext_type == ".csv":
                tables_gen = ext.extract_from_csv(fpath)
            else:
                continue

            for t in tables_gen:
                tables_count += 1
                for r in ext.map_and_clean([t]):
                    total_regs += 1

            if total_regs > 0:
                results["success"].append((rel, ext_type, tables_count, total_regs))
            else:
                results["zero_registers"].append((rel, ext_type, tables_count))
        except Exception as e:
            tb = traceback.format_exc()
            results["error"].append((rel, ext_type, str(e), tb))

    print(f"\n================ AUDIT SUMMARY ================")
    print(f"SUCCESS (Registers extracted): {len(results['success'])}")
    print(f"ZERO REGISTERS: {len(results['zero_registers'])}")
    print(f"ERRORS / EXCEPTIONS: {len(results['error'])}")

    if results["error"]:
        print(f"\n================ DETAILED ERRORS ================")
        for rel, ext_type, err, tb in results["error"]:
            print(f"\n[ERROR] {rel} ({ext_type}): {err}")
            print(tb[:500])

    if results["zero_registers"]:
        print(f"\n================ ZERO REGISTERS SAMPLE (First 20) ================")
        for rel, ext_type, tcount in results["zero_registers"][:20]:
            print(f"[ZERO] {rel} ({ext_type}) - tables found: {tcount}")


if __name__ == "__main__":
    audit_all()
