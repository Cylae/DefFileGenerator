#!/usr/bin/env python3
"""
Mass WebdynSunPM Definition Generator.
Generates the maximum number of compliant, validated Webdyn definition files
and stores them in 'generated_definitions/'.

Sources:
1. Existing generated_definitions (preserved and re-validated)
2. Local repository definitions (Definitions/*.csv) enriched with SunPM tags
3. Support network share (\\WDN-FILES.WEBDYN.COM\Support\Equipementiers):
   - Direct Webdyn definition CSVs (sanitized, deduplicated, enriched)
   - Vendor register table CSVs (extracted via Extractor -> Generator)
   - Vendor Excel files (extracted via Extractor -> Generator)
   - Key vendor PDF technical specs (extracted via Extractor -> Generator)
"""

import csv
import logging
import os
import re
import sys
from pathlib import Path

# Add project root to sys.path
WORKSPACE_DIR = Path(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
sys.path.insert(0, str(WORKSPACE_DIR))

from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator  # noqa: E402
from DefFileGenerator.extractor import Extractor, peek_generator  # noqa: E402
from DefFileGenerator.knowledge_miner import (  # noqa: E402
    auto_fix_definition_file,
    load_dataset_knowledge,
)

logging.basicConfig(level=logging.WARNING, format="%(levelname)s: %(message)s")

EQUIP_DIR = Path(r"\\WDN-FILES.WEBDYN.COM\Support\Equipementiers")
LOCAL_DEFS_DIR = WORKSPACE_DIR / "Definitions"
OUTPUT_DIR = WORKSPACE_DIR / "generated_definitions"
OUTPUT_DIR.mkdir(exist_ok=True)

generator = Generator()
extractor = Extractor()
knowledge = load_dataset_knowledge()

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
    if any(
        k in fn_upper
        for k in (
            "SUN2000",
            "INVERTER",
            "SOLAX",
            "GROWATT",
            "SMA",
            "KACO",
            "GOODWE",
            "FOX",
            "SOFAR",
            "SOLIS",
            "SUNGROW",
        )
    ):
        return "INV"
    if any(
        k in fn_upper
        for k in (
            "COUNTIS",
            "DIRIS",
            "EM210",
            "DTSU",
            "METER",
            "ANALYZER",
            "LETEL",
            "LETTEL",
            "JANITZA",
            "CIRCUTOR",
        )
    ):
        return "MET"
    if any(
        k in fn_upper
        for k in ("SR05", "SR20", "SR30", "DUSTIQ", "PYRA", "TEMP", "METEO", "WEATHER", "SENSOR")
    ):
        return "SEN"
    return "OTH"


print("=" * 80)
print(f"MASS DEFINITION GENERATOR TARGETING: {OUTPUT_DIR}")
print("=" * 80)

generated_files: set[str] = set()

# -------------------------------------------------------------------------
# Phase 1: Preserve and validate existing definitions in generated_definitions
# -------------------------------------------------------------------------
print("\n[Phase 1] Indexing and validating existing generated definitions...")
for existing in OUTPUT_DIR.glob("*.csv"):
    try:
        report = generator.validate_csv_detailed(str(existing), strict=False, strict_overlap=False)
        if report.is_valid and report.register_count > 0:
            generated_files.add(existing.name.lower())
    except Exception as e:
        print(f"  Warning on existing file {existing.name}: {e}")

print(f"Existing valid definitions preserved: {len(generated_files)}")

# -------------------------------------------------------------------------
# Phase 2: Process all Local Definitions in Definitions/*.csv
# -------------------------------------------------------------------------
print("\n[Phase 2] Processing local definitions from Definitions/ ...")
local_csvs = list(LOCAL_DEFS_DIR.glob("*.csv"))
added_local = 0
for l_csv in local_csvs:
    target_name = l_csv.name
    target_path = OUTPUT_DIR / target_name
    try:
        # Standardize and enrich with SunPM tags
        auto_fix_definition_file(
            input_path=l_csv,
            output_path=target_path,
            auto_tag=True,
            normalize_legacy=True,
            deduplicate=True,
        )
        report = generator.validate_csv_detailed(
            str(target_path), strict=False, strict_overlap=False
        )
        if report.is_valid and report.register_count > 0:
            generated_files.add(target_name.lower())
            added_local += 1
        else:
            # Revert if invalid
            if target_path.exists():
                target_path.unlink()
    except Exception as e:
        print(f"  Error on local {l_csv.name}: {e}")

print(f"Local definitions successfully generated & enriched: {added_local}")

# -------------------------------------------------------------------------
# Phase 3: Process Direct Webdyn CSV Definitions from Equipementiers share
# -------------------------------------------------------------------------
print("\n[Phase 3] Scanning and compiling direct Webdyn CSVs from Equipementiers share...")
share_csvs = [p for p in EQUIP_DIR.rglob("*.csv") if not p.name.startswith("~$")]
added_share_csv = 0

for s_csv in share_csvs:
    try:
        rel_parts = s_csv.relative_to(EQUIP_DIR).parts
        mfg = rel_parts[0] if len(rel_parts) > 1 else "Generic"

        # Read header to inspect conformance
        with open(s_csv, encoding="utf-8-sig", errors="ignore") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader, None)
            if not header or len(header) < 2:
                continue
            first_row = next(reader, None)
            if not first_row or len(first_row) < 7:
                continue

        # Check if first row has numeric register index and Modbus FC
        fc_val = first_row[1].strip()
        if fc_val not in ("1", "2", "3", "4", "01", "02", "03", "04"):
            continue

        raw_cat = header[1].strip() if len(header) > 1 else ""
        pfx = detect_prefix(raw_cat, s_csv.name, mfg)

        # Standardize target filename
        base_name = s_csv.stem
        # Strip redundant prefix if present
        for pref in ("INV_", "MET_", "SEN_", "COM_", "BAT_", "OTH_"):
            if base_name.upper().startswith(pref):
                base_name = base_name[len(pref) :]
                break

        # Avoid manufacturer duplication in filename
        if not base_name.lower().startswith(mfg.lower()):
            clean_mfg = sanitize_name(mfg)
            target_filename = f"{pfx}_{clean_mfg}_{sanitize_name(base_name)}.csv"
        else:
            target_filename = f"{pfx}_{sanitize_name(base_name)}.csv"

        if target_filename.lower() in generated_files:
            continue

        target_path = OUTPUT_DIR / target_filename

        # Clean and auto-tag
        counts = auto_fix_definition_file(
            input_path=s_csv,
            output_path=target_path,
            auto_tag=True,
            normalize_legacy=True,
            deduplicate=True,
        )

        # Validate
        report = generator.validate_csv_detailed(
            str(target_path), strict=False, strict_overlap=False
        )
        if report.is_valid and report.register_count >= 3:
            generated_files.add(target_filename.lower())
            added_share_csv += 1
        else:
            if target_path.exists():
                target_path.unlink()

    except Exception:
        pass

print(f"Share CSV definitions successfully compiled & validated: {added_share_csv}")

# -------------------------------------------------------------------------
# Phase 4: Process Vendor Excel Files (.xlsx) from Equipementiers share
# -------------------------------------------------------------------------
print("\n[Phase 4] Extracting & compiling Excel Modbus tables (.xlsx)...")
share_xlsx = [
    p
    for p in EQUIP_DIR.rglob("*.xlsx")
    if not p.name.startswith("~$") and p.stat().st_size < 30_000_000
]
added_xlsx = 0

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
        seen = set()
        dedup_rows = []
        for r in mapped:
            k = (str(r.get("RegisterType", "Holding")), str(r.get("Address", "")).strip())
            if k not in seen:
                seen.add(k)
                dedup_rows.append(r)

        target_path = OUTPUT_DIR / target_filename
        cfg = GeneratorConfig(
            input_file=str(x_file),
            output=str(target_path),
            manufacturer=mfg,
            model=base_name[:30],
            protocol="modbusRTU",
            category="Inverter"
            if pfx == "INV"
            else ("Meter" if pfx == "MET" else ("Sensor" if pfx == "SEN" else "Other")),
            address_offset=0,
            strict_validation=False,
        )

        ok = run_generator(cfg, input_data=dedup_rows)
        if ok and target_path.exists():
            report = generator.validate_csv_detailed(
                str(target_path), strict=False, strict_overlap=False
            )
            if report.is_valid and report.register_count >= 3:
                generated_files.add(target_filename.lower())
                added_xlsx += 1
            else:
                target_path.unlink()

    except Exception:
        pass

print(f"Excel files successfully compiled into valid definitions: {added_xlsx}")

# -------------------------------------------------------------------------
# Phase 5: Process Vendor CSV tables (non-Webdyn native CSVs)
# -------------------------------------------------------------------------
print("\n[Phase 5] Extracting & compiling Vendor Modbus CSV tables...")
added_vendor_csv = 0
for s_csv in share_csvs:
    try:
        rel_parts = s_csv.relative_to(EQUIP_DIR).parts
        mfg = rel_parts[0] if len(rel_parts) > 1 else "Generic"
        pfx = detect_prefix("", s_csv.name, mfg)
        base_name = s_csv.stem

        target_filename = f"{pfx}_{sanitize_name(mfg)}_{sanitize_name(base_name)}.csv"
        if target_filename.lower() in generated_files:
            continue

        raw = extractor.extract_from_csv(str(s_csv))
        has_data, peeked = peek_generator(raw)
        if not has_data:
            continue

        mapped = list(extractor.map_and_clean(peeked))
        if len(mapped) < 3:
            continue

        target_path = OUTPUT_DIR / target_filename
        cfg = GeneratorConfig(
            input_file=str(s_csv),
            output=str(target_path),
            manufacturer=mfg,
            model=base_name[:30],
            protocol="modbusRTU",
            category="Inverter"
            if pfx == "INV"
            else ("Meter" if pfx == "MET" else ("Sensor" if pfx == "SEN" else "Other")),
            address_offset=0,
            strict_validation=False,
        )

        ok = run_generator(cfg, input_data=mapped)
        if ok and target_path.exists():
            report = generator.validate_csv_detailed(
                str(target_path), strict=False, strict_overlap=False
            )
            if report.is_valid and report.register_count >= 3:
                generated_files.add(target_filename.lower())
                added_vendor_csv += 1
            else:
                target_path.unlink()

    except Exception:
        pass

print(f"Vendor CSV tables successfully converted: {added_vendor_csv}")

# -------------------------------------------------------------------------
# Final Summary
# -------------------------------------------------------------------------
all_final = list(OUTPUT_DIR.glob("*.csv"))
print("\n" + "=" * 80)
print(f"TOTAL DEFINITIONS GENERATED IN {OUTPUT_DIR}: {len(all_final)}")
print("=" * 80)
