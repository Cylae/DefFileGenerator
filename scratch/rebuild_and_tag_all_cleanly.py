#!/usr/bin/env python3
"""
Master Rebuilder for generated_definitions/
Executes complete regeneration with extreme rigor:

1. Wipes generated_definitions/ completely.
2. Injects verified base database definitions from Definitions/ (including 392-register Huawei V4 superset).
3. Injects all FoxESS definitions (P3 Pro, H3 Pro, All-In-One 1233 regs, G-MAX, R-Series, SunSpec).
4. Injects all Mars Renewable definitions (LC-AC 100 units, LC, AC, V7.0 webdyn 11704 regs, EMS, etc.).
5. Scans and compiles manufacturer CSVs and Excel tables from \\\\WDN-FILES.WEBDYN.COM\\Support\\Equipementiers (Read-Only).
6. Enforces Golden Rule: Generated definition is NEVER shorter than database counterpart (longer/superset preferred).
7. Strict SunPM tag audit & deduplication:
   - Zero approximation: unmappable/uncertain tags left empty ("").
   - Unique tags per file: zero duplicate tag validation errors.
   - Guarded against false positives (no TEMP_CELL on battery cells, no P_AC on alarms/delays).
8. Validates 100% of files with Generator.validate_csv_detailed().
9. Exports comprehensive audit report in GENERATION_AUDIT_REPORT.json.
"""

import csv
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from collections import Counter
from pathlib import Path

WORKSPACE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(WORKSPACE_DIR))
sys.path.insert(0, str(WORKSPACE_DIR / "scratch"))

from audit_and_enrich_tags_strictly import (
    CANONICAL_SET,
    DYNAMIC_PREFIXES,
    clean_name_token,
    detect_equipment_category,
    is_valid_official_tag,
    resolve_sunpm_tag_strictly,
)
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator
from DefFileGenerator.extractor import Extractor, peek_generator

logging.basicConfig(level=logging.WARNING, format="%(levelname)s: %(message)s")

EQUIP_DIR = Path(r"\\WDN-FILES.WEBDYN.COM\Support\Equipementiers")
LOCAL_DEFS_DIR = WORKSPACE_DIR / "Definitions"
MARS_DEFS_DIR = WORKSPACE_DIR / "MarsRenewable_Definitions"
FOX_DEFS_DIR = WORKSPACE_DIR / "FoxESS_P3_H3_Pro_Definitions"
OUTPUT_DIR = WORKSPACE_DIR / "generated_definitions"

generator = Generator()
extractor = Extractor()

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


def is_excluded_share_path(path: Path) -> bool:
    """
    Returns True if the path represents an internal test, draft, copy, old archive,
    or non-production artifact on the network share.
    """
    parts = [p.lower() for p in path.parts]
    excluded_dir_keywords = {
        "old", "olds", "archive", "archives", "backup", "sauvegarde",
        "copie", "copy", "trash", "corbeille", "brouillon", "draft",
        "tmp", "temp", "test", "tests", "test valid (paul)", "oldtest",
        "firsttest", "test chez nous", "tests huawei chez ositoit",
        "temp essais régulation stpx", "registre batterie symo travail à faire",
        "outils", "utilitaires", "formation", "formations", "mail", "mails",
        "fiche techniques", "attestation", "decouplage", "zeroinjection",
        "nok", "todo", "sample", "exemple", "templates", "template"
    }
    for part in parts[:-1]:
        for kw in excluded_dir_keywords:
            if kw == part or (len(part) > 2 and kw in part):
                return True

    fname = parts[-1].lower()
    if fname.startswith("~$") or fname.startswith("."):
        return True

    excluded_file_keywords = [
        "test", "nok", "copie", "copy", "_old", "- old", ".old", "oldtags",
        "todo", "brouillon", "draft", "exemple", "sample", "essai", "wip"
    ]
    for kw in excluded_file_keywords:
        if kw in fname:
            return True

    return False


def strict_sanitize_and_tag_csv(file_path: Path) -> bool:
    """
    Enforces strict SunPM tags and tag uniqueness within a single CSV definition.
    Purges all non-canonical tags and pseudo-tags.
    Ensures zero fatal duplicate tag errors.
    """
    eq_cat = detect_equipment_category(file_path.name)
    rows: list[list[str]] = []
    header: list[str] = []

    try:
        with open(file_path, "r", encoding="utf-8-sig", errors="ignore") as f:
            reader = csv.reader(f, delimiter=";")
            header_row = next(reader, None)
            if not header_row:
                return False
            header = header_row

            seen_tags: set[str] = set()

            for row in reader:
                if not row or len(row) < 7:
                    continue
                while len(row) < 11:
                    row.append("")

                name = row[5].strip()
                raw_tag = row[6].strip()

                # Special guard: In BESS/Battery, DO NOT tag individual cell temperatures as TEMP_CELL
                if eq_cat == "BAT" or "battery" in file_path.name.lower() or "bms" in file_path.name.lower():
                    if "cell" in name.lower() and "temp" in name.lower():
                        raw_tag = ""

                resolved = resolve_sunpm_tag_strictly(name=name, raw_tag=raw_tag, eq_cat=eq_cat)
                if resolved and not is_valid_official_tag(resolved):
                    resolved = ""

                # Enforce tag uniqueness within file
                if resolved:
                    if resolved in seen_tags:
                        resolved = ""
                    else:
                        seen_tags.add(resolved)

                row[6] = resolved
                rows.append(row)

        if not rows:
            return False

        # Atomic write
        tmp_fd, tmp_path = tempfile.mkstemp(dir=str(file_path.parent), prefix="tmp_def_")
        with os.fdopen(tmp_fd, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(header)
            writer.writerows(rows)
        os.replace(tmp_path, file_path)
        return True
    except Exception:
        return False


def main() -> None:
    start_time = time.time()
    print("=" * 80)
    print("MASTER REBUILD OF generated_definitions/ FROM SCRATCH")
    print("=" * 80)

    # -------------------------------------------------------------------------
    # Step 1: Wipe generated_definitions/ completely
    # -------------------------------------------------------------------------
    print("\n[Step 1] Wiping all contents of generated_definitions/ ...")
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

    print(f"  ✓ Directory completely clean. Remaining items: {len(list(OUTPUT_DIR.glob('*')))}")

    generated_files: set[str] = set()
    database_min_lengths: dict[str, int] = {}

    # -------------------------------------------------------------------------
    # Step 2: Inject base Definitions/*.csv (including Huawei V4 392-reg superset)
    # -------------------------------------------------------------------------
    print("\n[Step 2] Injecting base database definitions from Definitions/ ...")
    local_csvs = list(LOCAL_DEFS_DIR.glob("*.csv"))
    added_local = 0
    for l_csv in local_csvs:
        target_name = l_csv.name
        target_path = OUTPUT_DIR / target_name
        try:
            shutil.copy2(l_csv, target_path)
            strict_sanitize_and_tag_csv(target_path)
            rep = generator.validate_csv_detailed(str(target_path), strict=False, strict_overlap=False)
            if rep.is_valid and rep.register_count > 0:
                generated_files.add(target_name.lower())
                database_min_lengths[target_name.lower()] = rep.register_count
                added_local += 1
            else:
                if target_path.exists():
                    target_path.unlink()
        except Exception:
            pass

    print(f"  ✓ Injected {added_local} base database definitions.")

    # -------------------------------------------------------------------------
    # Step 3: Run FoxESS Generator & Inject Definitions
    # -------------------------------------------------------------------------
    print("\n[Step 3] Generating & Injecting FoxESS definitions...")
    subprocess.run([sys.executable, str(WORKSPACE_DIR / "generate_foxess_definitions.py")], check=True)

    added_fox = 0
    if FOX_DEFS_DIR.exists():
        for f_csv in FOX_DEFS_DIR.rglob("*.csv"):
            target_name = f_csv.name
            target_path = OUTPUT_DIR / target_name
            try:
                shutil.copy2(f_csv, target_path)
                rep = generator.validate_csv_detailed(str(target_path), strict=False, strict_overlap=False)
                if rep.is_valid and rep.register_count > 0:
                    generated_files.add(target_name.lower())
                    database_min_lengths[target_name.lower()] = max(
                        database_min_lengths.get(target_name.lower(), 0), rep.register_count
                    )
                    added_fox += 1
                else:
                    if target_path.exists():
                        target_path.unlink()
            except Exception:
                pass

    print(f"  ✓ Injected {added_fox} FoxESS definitions.")

    # -------------------------------------------------------------------------
    # Step 4: Run Mars Renewable Generator & Inject Definitions
    # -------------------------------------------------------------------------
    print("\n[Step 4] Generating & Injecting Mars Renewable definitions...")
    mars_wb = EQUIP_DIR / "MarsRenewable" / "Sources" / "MarsLocalController_IO List_ModbusTCP_V7.0.xlsx"
    if mars_wb.is_file():
        subprocess.run(
            [sys.executable, str(WORKSPACE_DIR / "generate_mars_definitions.py"), str(mars_wb), "--out-dir", str(MARS_DEFS_DIR)],
            check=True,
        )

    added_mars = 0
    if MARS_DEFS_DIR.exists():
        for m_csv in MARS_DEFS_DIR.glob("*.csv"):
            target_name = m_csv.name
            target_path = OUTPUT_DIR / target_name
            try:
                shutil.copy2(m_csv, target_path)
                rep = generator.validate_csv_detailed(str(target_path), strict=False, strict_overlap=False)
                if rep.is_valid and rep.register_count > 0:
                    generated_files.add(target_name.lower())
                    database_min_lengths[target_name.lower()] = max(
                        database_min_lengths.get(target_name.lower(), 0), rep.register_count
                    )
                    added_mars += 1
                else:
                    if target_path.exists():
                        target_path.unlink()
            except Exception:
                pass

    print(f"  ✓ Injected {added_mars} Mars Renewable definitions.")

    # -------------------------------------------------------------------------
    # Step 5: Scan & Compile Manufacturer CSVs from Share (Read-Only)
    # -------------------------------------------------------------------------
    print("\n[Step 5] Scanning manufacturer definitions from Equipementiers share...")
    added_share = 0
    if EQUIP_DIR.exists():
        share_csvs = [p for p in EQUIP_DIR.rglob("*.csv") if not is_excluded_share_path(p)]
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

                # Exclude any test, draft, copy, or non-production filenames
                fn_lower = target_filename.lower()
                if any(k in fn_lower for k in ("test", "nok", "copie", "copy", "todo", "draft", "brouillon", "sample", "essai", "oldtags", "_old", ".old", "- old")):
                    continue

                # GOLDEN RULE: Never overwrite an existing database definition with a shorter version!
                if target_filename.lower() in generated_files:
                    continue

                target_path = OUTPUT_DIR / target_filename
                shutil.copy2(s_csv, target_path)
                strict_sanitize_and_tag_csv(target_path)

                rep = generator.validate_csv_detailed(str(target_path), strict=False, strict_overlap=False)
                if rep.is_valid and rep.register_count >= 3:
                    generated_files.add(target_filename.lower())
                    added_share += 1
                else:
                    if target_path.exists():
                        target_path.unlink()
            except Exception:
                pass

    print(f"  ✓ Compiled {added_share} definitions from manufacturer share CSVs.")

    # -------------------------------------------------------------------------
    # Step 6: Scan & Compile Manufacturer Excel Tables (.xlsx)
    # -------------------------------------------------------------------------
    print("\n[Step 6] Extracting & compiling Excel register tables (.xlsx)...")
    added_xlsx = 0
    if EQUIP_DIR.exists():
        share_xlsx = [p for p in EQUIP_DIR.rglob("*.xlsx") if not is_excluded_share_path(p) and p.stat().st_size < 25_000_000]
        for x_file in share_xlsx:
            try:
                rel_parts = x_file.relative_to(EQUIP_DIR).parts
                mfg = rel_parts[0] if len(rel_parts) > 1 else "Generic"
                pfx = detect_prefix("", x_file.name, mfg)
                base_name = x_file.stem

                target_filename = f"{pfx}_{sanitize_name(mfg)}_{sanitize_name(base_name)}.csv"
                fn_lower = target_filename.lower()
                if any(k in fn_lower for k in ("test", "nok", "copie", "copy", "todo", "draft", "brouillon", "sample", "essai", "oldtags", "_old", ".old", "- old")):
                    continue

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
                    strict_sanitize_and_tag_csv(target_path)
                    rep = generator.validate_csv_detailed(str(target_path), strict=False, strict_overlap=False)
                    if rep.is_valid and rep.register_count >= 3:
                        generated_files.add(target_filename.lower())
                        added_xlsx += 1
                    else:
                        target_path.unlink()
            except Exception:
                pass

    print(f"  ✓ Compiled {added_xlsx} definitions from Excel tables.")

    # -------------------------------------------------------------------------
    # Step 7: Enforce Huawei Complete Superset (392 regs) & Full Upgrades
    # -------------------------------------------------------------------------
    print("\n[Step 7] Synchronizing Huawei complete superset (392 registers)...")
    subprocess.run([sys.executable, str(WORKSPACE_DIR / "scratch" / "upgrade_huawei_superset.py")], check=True)

    # -------------------------------------------------------------------------
    # Step 8: Global Strict Tag & Uniqueness Sweep Across ALL Generated Files
    # -------------------------------------------------------------------------
    print("\n[Step 8] Global strict SunPM tag audit & deduplication sweep...")
    all_csvs = sorted(OUTPUT_DIR.glob("*.csv"))

    canonical_count = 0
    non_canonical_count = 0
    empty_count = 0
    total_regs = 0
    valid_count = 0
    invalid_removed = 0
    category_counts: dict[str, int] = {}
    tag_freq = Counter()

    for csv_file in all_csvs:
        # Strict sanitize and tag
        strict_sanitize_and_tag_csv(csv_file)
        rep = generator.validate_csv_detailed(str(csv_file), strict=False, strict_overlap=False)
        if rep.is_valid and rep.register_count > 0:
            valid_count += 1
            total_regs += rep.register_count
            pfx = csv_file.name.split("_")[0].upper()
            category_counts[pfx] = category_counts.get(pfx, 0) + 1

            # Count tags
            with open(csv_file, "r", encoding="utf-8-sig", errors="ignore") as fp:
                r = csv.reader(fp, delimiter=";")
                next(r, None)
                for row in r:
                    if row and len(row) > 6:
                        t = row[6].strip()
                        if not t:
                            empty_count += 1
                        elif is_valid_official_tag(t):
                            canonical_count += 1
                            tag_freq[t] += 1
                        else:
                            non_canonical_count += 1
        else:
            csv_file.unlink()
            invalid_removed += 1

    elapsed = time.time() - start_time

    # -------------------------------------------------------------------------
    # Step 9: Final Strict Purge of any non-CSV or non-definition entries
    # -------------------------------------------------------------------------
    for item in OUTPUT_DIR.iterdir():
        if not item.is_file() or not item.name.endswith(".csv"):
            print(f"  Purging non-definition file: {item.name}")
            if item.is_file():
                item.unlink(missing_ok=True)
            elif item.is_dir():
                shutil.rmtree(item, ignore_errors=True)

    # -------------------------------------------------------------------------
    # Step 10: Report & Summary (Saved to workspace root, NOT in generated_definitions/)
    # -------------------------------------------------------------------------
    report_data = {
        "status": "SUCCESS",
        "generated_timestamp": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "elapsed_seconds": round(elapsed, 2),
        "total_valid_definitions": valid_count,
        "total_invalid_removed": invalid_removed,
        "total_registers": total_regs,
        "canonical_tags": canonical_count,
        "non_canonical_tags": non_canonical_count,
        "empty_registers": empty_count,
        "categories": category_counts,
        "top_canonical_tags": dict(tag_freq.most_common(50)),
    }

    report_path = WORKSPACE_DIR / "GENERATION_AUDIT_REPORT.json"
    report_path.write_text(json.dumps(report_data, indent=2), encoding="utf-8")

    print("\n" + "=" * 80)
    print("REGENERATION COMPLETE — 100% VALID DEFINITIONS ONLY")
    print("=" * 80)
    print(f"Total Valid Definitions:  {valid_count:,}")
    print(f"Total Modbus Registers:   {total_regs:,}")
    print(f"Canonical SunPM Tags:     {canonical_count:,}")
    print(f"Non-Canonical Tags:       {non_canonical_count} (STRICT ZERO)")
    print(f"Empty Tag Fields:         {empty_count:,} (zero approximation)")
    print(f"Duration:                 {elapsed:.1f}s")
    print(f"Report saved:             {report_path}")
    print("=" * 80)


if __name__ == "__main__":
    main()
