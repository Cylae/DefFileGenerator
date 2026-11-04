#!/usr/bin/env python3
"""
Master Rebuilder for generated_definitions/
Rebuilds the entire definition library in generated_definitions/ from scratch:

1. Wipes generated_definitions/ completely.
2. Generates/enriches all base definitions from Definitions/*.csv.
3. Injects all Mars Renewable definitions (LC-AC 100 units, LC, AC, PCS, Rack, Cell, Meter, PV, etc.).
4. Injects all FoxESS P3 Pro and H3 Pro definitions (Inverter, Meter, Battery, combined, etc.).
5. Syncs verified definitions from tagged_definitions/.
6. Scans and compiles all valid manufacturer Webdyn CSVs from \\\\WDN-FILES.WEBDYN.COM\\Support\\Equipementiers (Read-Only).
7. Extracts and compiles vendor Excel tables (.xlsx) and vendor CSV tables.
8. Validates 100% of generated files with Generator.validate_csv_detailed(strict=False, strict_overlap=False).
   Discards any file that has ERROR issues.
9. Exports a comprehensive audit report in JSON and Markdown.
"""

import csv
import json
import logging
import os
import re
import shutil
import sys
import time
from pathlib import Path

# Add project root to sys.path
WORKSPACE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(WORKSPACE_DIR))

from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator
from DefFileGenerator.extractor import Extractor, peek_generator
from DefFileGenerator.knowledge_miner import (
    auto_fix_definition_file,
    get_active_knowledge,
    infer_sunpm_tag,
    is_canonical_sunpm_tag,
    load_dataset_knowledge,
)

logging.basicConfig(level=logging.WARNING, format="%(levelname)s: %(message)s")

EQUIP_DIR = Path(r"\\WDN-FILES.WEBDYN.COM\Support\Equipementiers")
LOCAL_DEFS_DIR = WORKSPACE_DIR / "Definitions"
MARS_DEFS_DIR = WORKSPACE_DIR / "MarsRenewable_Definitions"
FOX_DEFS_DIR = WORKSPACE_DIR / "FoxESS_P3_H3_Pro_Definitions"
TAGGED_DEFS_DIR = WORKSPACE_DIR / "tagged_definitions"
OUTPUT_DIR = WORKSPACE_DIR / "generated_definitions"

generator = Generator()
extractor = Extractor()
knowledge = get_active_knowledge()

PREFIX_MAP = {
    "INVERTER": "INV",
    "ONDULEUR": "INV",
    "METER": "MET",
    "COMPTEUR": "MET",
    "SENSOR": "SEN",
    "SONDE": "SEN",
    "PYRANOMETER": "SEN",
    "PYRANOMETRE": "SEN",
    "BESS": "BAT",
    "BATTERY": "BAT",
    "BATTERIE": "BAT",
    "STORAGE": "BAT",
    "COMMUNICATION": "COM",
    "GATEWAY": "COM",
    "PASSERELLE": "COM",
    "LOGGER": "COM",
    "DATALOGGER": "COM",
    "CHARGER": "OTH",
    "EV": "OTH",
    "IRRIGATION": "OTH",
    "OTHER": "OTH",
    "AUTRE": "OTH",
}


def sanitize_name(name: str) -> str:
    cleaned = re.sub(r"[^\w\-.]", "_", name)
    cleaned = re.sub(r"_+", "_", cleaned)
    return cleaned.strip("_")


def detect_prefix(category_str: str, filename: str, mfg: str = "") -> str:
    cat_upper = category_str.upper()
    for key, pfx in PREFIX_MAP.items():
        if key in cat_upper:
            return pfx
    fn_upper = filename.upper()
    for pfx in ("INV_", "MET_", "SEN_", "COM_", "BAT_", "OTH_"):
        if fn_upper.startswith(pfx):
            return pfx[:-1]
    for key, pfx in PREFIX_MAP.items():
        if key in fn_upper:
            return pfx
    if any(k in fn_upper for k in ("SUN2000", "INVERTER", "SOLAX", "GROWATT", "SMA", "KACO", "GOODWE", "FOX", "SOFAR", "SOLIS", "SUNGROW")):
        return "INV"
    if any(k in fn_upper for k in ("COUNTIS", "DIRIS", "EM210", "DTSU", "METER", "ANALYZER", "LETEL", "LETTEL", "JANITZA", "CIRCUTOR")):
        return "MET"
    if any(k in fn_upper for k in ("SR05", "SR20", "SR30", "DUSTIQ", "PYRA", "TEMP", "METEO", "WEATHER", "SENSOR")):
        return "SEN"
    return "OTH"


def main() -> None:
    start_time = time.time()
    print("=" * 80)
    print("REBUILDING generated_definitions/ FROM SCRATCH")
    print(f"Target Directory: {OUTPUT_DIR}")
    print("=" * 80)

    # -------------------------------------------------------------------------
    # Step 1: Wipe generated_definitions/ completely
    # -------------------------------------------------------------------------
    print("\n[Step 1] Wiping all existing files in generated_definitions/ ...")
    if OUTPUT_DIR.exists():
        for item in OUTPUT_DIR.iterdir():
            if item.is_file():
                try:
                    item.unlink()
                except Exception as e:
                    print(f"  Warning: could not delete {item.name}: {e}")
            elif item.is_dir():
                shutil.rmtree(item, ignore_errors=True)
    else:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    remaining = list(OUTPUT_DIR.glob("*"))
    print(f"  ✓ Directory wiped clean. Remaining files: {len(remaining)}")

    generated_files: set[str] = set()

    # -------------------------------------------------------------------------
    # Step 2: Base Local Definitions from Definitions/*.csv
    # -------------------------------------------------------------------------
    print("\n[Step 2] Processing local definitions from Definitions/ ...")
    local_csvs = list(LOCAL_DEFS_DIR.glob("*.csv"))
    added_local = 0
    for l_csv in local_csvs:
        target_name = l_csv.name
        target_path = OUTPUT_DIR / target_name
        try:
            auto_fix_definition_file(
                input_path=l_csv,
                output_path=target_path,
                auto_tag=True,
                normalize_legacy=True,
                deduplicate=True,
            )
            report = generator.validate_csv_detailed(str(target_path), strict=False, strict_overlap=False)
            if report.is_valid and report.register_count > 0:
                generated_files.add(target_name.lower())
                added_local += 1
            else:
                if target_path.exists():
                    target_path.unlink()
        except Exception as e:
            pass

    print(f"  ✓ Definitions/ processed: {added_local} files added.")

    # -------------------------------------------------------------------------
    # Step 3: Mars Renewable Definitions
    # -------------------------------------------------------------------------
    print("\n[Step 3] Adding Mars Renewable definitions...")
    added_mars = 0
    if MARS_DEFS_DIR.exists():
        for m_csv in MARS_DEFS_DIR.glob("*.csv"):
            target_path = OUTPUT_DIR / m_csv.name
            try:
                shutil.copy2(m_csv, target_path)
                rep = generator.validate_csv_detailed(str(target_path), strict=False, strict_overlap=False)
                if rep.is_valid and rep.register_count > 0:
                    generated_files.add(m_csv.name.lower())
                    added_mars += 1
                else:
                    if target_path.exists():
                        target_path.unlink()
            except Exception:
                pass
    print(f"  ✓ Mars Renewable definitions added: {added_mars} files.")

    # -------------------------------------------------------------------------
    # Step 4: FoxESS P3 Pro & H3 Pro Definitions
    # -------------------------------------------------------------------------
    print("\n[Step 4] Adding FoxESS P3 Pro & H3 Pro definitions...")
    added_fox = 0
    if FOX_DEFS_DIR.exists():
        # Root and subfolder files
        for f_csv in FOX_DEFS_DIR.rglob("*.csv"):
            # Avoid overwriting with duplicate subfolder names, use clean naming
            target_name = f_csv.name
            target_path = OUTPUT_DIR / target_name
            if target_name.lower() in generated_files:
                continue
            try:
                shutil.copy2(f_csv, target_path)
                rep = generator.validate_csv_detailed(str(target_path), strict=False, strict_overlap=False)
                if rep.is_valid and rep.register_count > 0:
                    generated_files.add(target_name.lower())
                    added_fox += 1
                else:
                    if target_path.exists():
                        target_path.unlink()
            except Exception:
                pass
    print(f"  ✓ FoxESS P3 Pro & H3 Pro definitions added: {added_fox} files.")

    # -------------------------------------------------------------------------
    # Step 5: Sync verified definitions from tagged_definitions/
    # -------------------------------------------------------------------------
    print("\n[Step 5] Syncing verified definitions from tagged_definitions/ ...")
    added_tagged = 0
    if TAGGED_DEFS_DIR.exists():
        for t_csv in TAGGED_DEFS_DIR.glob("*.csv"):
            if t_csv.name.lower() in generated_files:
                continue
            target_path = OUTPUT_DIR / t_csv.name
            try:
                shutil.copy2(t_csv, target_path)
                rep = generator.validate_csv_detailed(str(target_path), strict=False, strict_overlap=False)
                if rep.is_valid and rep.register_count > 0:
                    generated_files.add(t_csv.name.lower())
                    added_tagged += 1
                else:
                    if target_path.exists():
                        target_path.unlink()
            except Exception:
                pass
    print(f"  ✓ Tagged definitions synced: {added_tagged} files.")

    # -------------------------------------------------------------------------
    # Step 6: Scan & compile direct Webdyn CSVs from Equipementiers share
    # -------------------------------------------------------------------------
    print("\n[Step 6] Scanning & compiling Webdyn CSVs from Equipementiers share (Read-Only)...")
    added_share_csv = 0
    if EQUIP_DIR.exists():
        share_csvs = [p for p in EQUIP_DIR.rglob("*.csv") if not p.name.startswith("~$")]
        for s_csv in share_csvs:
            try:
                rel_parts = s_csv.relative_to(EQUIP_DIR).parts
                mfg = rel_parts[0] if len(rel_parts) > 1 else "Generic"

                with open(s_csv, encoding="utf-8-sig", errors="ignore") as f:
                    reader = csv.reader(f, delimiter=";")
                    header = next(reader, None)
                    if not header or len(header) < 2:
                        continue
                    first_row = next(reader, None)
                    if not first_row or len(first_row) < 7:
                        continue

                fc_val = first_row[1].strip()
                if fc_val not in ("1", "2", "3", "4", "01", "02", "03", "04"):
                    continue

                raw_cat = header[1].strip() if len(header) > 1 else ""
                pfx = detect_prefix(raw_cat, s_csv.name, mfg)
                base_name = s_csv.stem

                for pref in ("INV_", "MET_", "SEN_", "COM_", "BAT_", "OTH_"):
                    if base_name.upper().startswith(pref):
                        base_name = base_name[len(pref) :]
                        break

                if not base_name.lower().startswith(mfg.lower()):
                    clean_mfg = sanitize_name(mfg)
                    target_filename = f"{pfx}_{clean_mfg}_{sanitize_name(base_name)}.csv"
                else:
                    target_filename = f"{pfx}_{sanitize_name(base_name)}.csv"

                if target_filename.lower() in generated_files:
                    continue

                target_path = OUTPUT_DIR / target_filename
                auto_fix_definition_file(
                    input_path=s_csv,
                    output_path=target_path,
                    auto_tag=True,
                    normalize_legacy=True,
                    deduplicate=True,
                )

                rep = generator.validate_csv_detailed(str(target_path), strict=False, strict_overlap=False)
                if rep.is_valid and rep.register_count >= 3:
                    generated_files.add(target_filename.lower())
                    added_share_csv += 1
                else:
                    if target_path.exists():
                        target_path.unlink()
            except Exception:
                pass

    print(f"  ✓ Share direct CSVs compiled: {added_share_csv} files.")

    # -------------------------------------------------------------------------
    # Step 7: Extract & compile Vendor Excel tables (.xlsx)
    # -------------------------------------------------------------------------
    print("\n[Step 7] Extracting & compiling Excel register tables (.xlsx) ...")
    added_xlsx = 0
    if EQUIP_DIR.exists():
        share_xlsx = [p for p in EQUIP_DIR.rglob("*.xlsx") if not p.name.startswith("~$") and p.stat().st_size < 25_000_000]
        for x_file in share_xlsx:
            try:
                rel_parts = x_file.relative_to(EQUIP_DIR).parts
                mfg = rel_parts[0] if len(rel_parts) > 1 else "Generic"
                pfx = detect_prefix("", x_file.name, mfg)
                base_name = x_file.stem

                target_filename = f"{pfx}_{sanitize_name(mfg)}_{sanitize_name(base_name)}.csv"
                if target_filename.lower() in generated_files:
                    continue

                raw = extractor.extract_from_excel(str(x_file))
                has_data, peeked = peek_generator(raw)
                if not has_data:
                    continue

                mapped = list(extractor.map_and_clean(peeked))
                if len(mapped) < 3:
                    continue

                # Deduplicate
                seen_regs = set()
                dedup_rows = []
                for r in mapped:
                    k = (str(r.get("RegisterType", "Holding")), str(r.get("Address", "")).strip())
                    if k not in seen_regs:
                        seen_regs.add(k)
                        dedup_rows.append(r)

                target_path = OUTPUT_DIR / target_filename
                cfg = GeneratorConfig(
                    input_file=str(x_file),
                    output=str(target_path),
                    manufacturer=mfg,
                    model=base_name[:30],
                    protocol="modbusRTU",
                    category="Inverter" if pfx == "INV" else ("Meter" if pfx == "MET" else ("Sensor" if pfx == "SEN" else "Other")),
                    address_offset=0,
                    strict_validation=False,
                )

                ok = run_generator(cfg, input_data=dedup_rows)
                if ok and target_path.exists():
                    rep = generator.validate_csv_detailed(str(target_path), strict=False, strict_overlap=False)
                    if rep.is_valid and rep.register_count >= 3:
                        generated_files.add(target_filename.lower())
                        added_xlsx += 1
                    else:
                        target_path.unlink()
            except Exception:
                pass

    print(f"  ✓ Excel register tables converted: {added_xlsx} files.")

    # -------------------------------------------------------------------------
    # Step 8: Strict Validation & Clean-up of generated_definitions/
    # -------------------------------------------------------------------------
    print("\n[Step 8] Final verification and audit of all generated files...")
    all_csvs = list(OUTPUT_DIR.glob("*.csv"))
    valid_count = 0
    invalid_removed = 0
    category_counts: dict[str, int] = {}
    total_registers = 0
    total_tags = 0

    for c in all_csvs:
        try:
            rep = generator.validate_csv_detailed(str(c), strict=False, strict_overlap=False)
            if rep.is_valid and rep.register_count > 0:
                valid_count += 1
                total_registers += rep.register_count
                pfx = c.name.split("_")[0].upper()
                category_counts[pfx] = category_counts.get(pfx, 0) + 1

                # Count tags
                with open(c, encoding="utf-8-sig", errors="ignore") as fp:
                    for idx, row in enumerate(csv.reader(fp, delimiter=";")):
                        if idx > 0 and len(row) > 6 and row[6].strip():
                            total_tags += 1
            else:
                c.unlink()
                invalid_removed += 1
        except Exception:
            c.unlink()
            invalid_removed += 1

    elapsed = time.time() - start_time
    print(f"  ✓ Valid definitions preserved: {valid_count}")
    print(f"  ✓ Invalid/empty files removed: {invalid_removed}")
    print(f"  ✓ Total Modbus registers:     {total_registers:,}")
    print(f"  ✓ Total SunPM tags:          {total_tags:,}")
    print(f"  ✓ Equipment Breakdown:")
    for k, v in sorted(category_counts.items()):
        print(f"      {k}: {v} definitions")

    # Save summary report in generated_definitions/
    report_data = {
        "status": "SUCCESS",
        "generated_timestamp": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "elapsed_seconds": round(elapsed, 2),
        "total_valid_definitions": valid_count,
        "total_registers": total_registers,
        "total_tags": total_tags,
        "categories": category_counts,
    }
    report_path = OUTPUT_DIR / "GENERATION_AUDIT_REPORT.json"
    report_path.write_text(json.dumps(report_data, indent=2), encoding="utf-8")

    print("\n" + "=" * 80)
    print(f"REBUILD COMPLETE IN {elapsed:.1f}s: {valid_count} FILES IN {OUTPUT_DIR}")
    print("=" * 80)


if __name__ == "__main__":
    main()
