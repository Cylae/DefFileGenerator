#!/usr/bin/env python3
"""
Mass Tagged Definition Generator.
Generates fully Webdyn-compliant definition files with SunPM tags
into a dedicated directory: 'tagged_definitions/'.

Source documents are accessed strictly READ-ONLY (no source files modified).
Leverages trained knowledge base with 3,098 mappings and official SunPM tags.
"""

import csv
import logging
import os
import re
import sys
import time
from pathlib import Path

# Add project root
WORKSPACE_DIR = Path(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
sys.path.insert(0, str(WORKSPACE_DIR))

from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator  # noqa: E402
from DefFileGenerator.extractor import Extractor, peek_generator  # noqa: E402
from DefFileGenerator.knowledge_miner import (  # noqa: E402
    clean_variable_key,
    get_active_knowledge,
    infer_sunpm_tag,
    is_canonical_sunpm_tag,
    load_dataset_knowledge,
    normalize_tag_to_sunpm,
)

logging.basicConfig(level=logging.WARNING, format="%(levelname)s: %(message)s")

EQUIP_DIR = Path(r"\\WDN-FILES.WEBDYN.COM\Support\Equipementiers")
LOCAL_DEFS_DIR = WORKSPACE_DIR / "Definitions"
OUTPUT_DIR = WORKSPACE_DIR / "tagged_definitions"
OUTPUT_DIR.mkdir(exist_ok=True)

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
        for k in (
            "SR05",
            "SR20",
            "SR30",
            "DUSTIQ",
            "PYRA",
            "TEMP",
            "METEO",
            "WEATHER",
            "SENSOR",
        )
    ):
        return "SEN"
    if any(k in fn_upper for k in ("BMS", "BATTERY", "BESS", "STORAGE", "RACK")):
        return "BAT"
    return "OTH"


def build_and_tag_rows(
    raw_rows: list[list[str]], detected_type: str = "Inverter"
) -> list[list[str]]:
    """Takes raw rows and formats them into standard 11-column rows with unique SunPM tags."""
    seen_rows = set()
    seen_tags = set()
    clean_rows = []

    for r in raw_rows:
        if len(r) < 6:
            continue
        fc = r[1].strip() if len(r) > 1 and r[1].strip() else "3"
        addr = r[2].strip() if len(r) > 2 else ""
        dtype = r[3].strip().upper() if len(r) > 3 and r[3].strip() else "U16"
        var_name = r[5].strip() if len(r) > 5 else ""
        existing_tag = r[6].strip() if len(r) > 6 else ""

        if not addr or not var_name:
            continue

        row_key = (fc, addr, dtype, var_name.lower())
        if row_key in seen_rows:
            continue
        seen_rows.add(row_key)

        # Tag determination
        target_tag = ""
        if existing_tag:
            norm = normalize_tag_to_sunpm(existing_tag)
            if norm and norm not in seen_tags:
                target_tag = norm
                seen_tags.add(target_tag)
        if not target_tag:
            inferred = infer_sunpm_tag(
                variable_name=var_name,
                existing_tag="",
                equipment_type=detected_type,
                knowledge=knowledge,
            )
            if inferred and inferred not in seen_tags:
                target_tag = inferred
                seen_tags.add(inferred)

        # Coef formatting
        raw_ca = r[7].strip() if len(r) > 7 else "1.0"
        raw_cb = r[8].strip() if len(r) > 8 else "0.0"
        try:
            _ = float(raw_ca)
            ca_str = raw_ca if raw_ca else "1.0"
        except ValueError:
            ca_str = "1.0"
        try:
            _ = float(raw_cb)
            cb_str = raw_cb if raw_cb else "0.0"
        except ValueError:
            cb_str = "0.0"

        unit = r[9].strip() if len(r) > 9 else ""
        raw_act = r[10].strip() if len(r) > 10 else ""
        raw_act_upper = raw_act.upper()
        if raw_act_upper in ("RO", "R", "READ", "READ-ONLY", "READ ONLY"):
            action = "4"
        elif raw_act_upper in ("RW", "W", "WRITE", "READ/WRITE", "READ-WRITE", "R/W", "WO", "WRITE-ONLY", "WRITE ONLY"):
            action = "1"
        elif raw_act in ("0", "1", "2", "4", "6", "7", "8", "9"):
            action = raw_act
        else:
            action = ""

        if not action:
            if target_tag.startswith("CMD_") or target_tag.endswith("_ENA"):
                action = "1"
            elif target_tag in ("STATE", "ALARM") or "alarm" in var_name.lower():
                action = "9"
            else:
                action = "4"
        elif target_tag.startswith("CMD_") and action not in ("1", "0"):
            action = "1"
        elif (target_tag.startswith("STATE") or target_tag.startswith("ALARM")) and action == "4":
            action = "9"

        clean_row = [
            str(len(clean_rows) + 1),
            fc,
            addr,
            dtype,
            "",  # ScaleFactor
            var_name,
            target_tag,
            ca_str,
            cb_str,
            unit,
            action,
        ]
        clean_rows.append(clean_row)

    return clean_rows


print("=" * 80)
print(f"GENERATING FULLY TAGGED DEFINITIONS IN: {OUTPUT_DIR}")
print("Constraint: Sources are strictly READ-ONLY.")
print("=" * 80)

generated_files: set[str] = set()
total_tags_assigned = 0

# -------------------------------------------------------------------------
# Phase 1: Local core definitions from Definitions/
# -------------------------------------------------------------------------
print("\n[Phase 1] Processing local definitions from Definitions/ ...")
local_csvs = sorted(LOCAL_DEFS_DIR.glob("*.csv"))
added_local = 0

for l_csv in local_csvs:
    target_name = l_csv.name
    target_path = OUTPUT_DIR / target_name
    try:
        # Strictly open read-only
        with open(l_csv, "r", encoding="utf-8-sig", errors="ignore") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader, None)
            if not header:
                continue
            rows = list(reader)

        eq_type = header[1] if len(header) > 1 and header[1].strip() else "Inverter"
        clean_rows = build_and_tag_rows(rows, detected_type=eq_type)
        if len(clean_rows) < 3:
            continue

        header11 = (header + [""] * 11)[:11]

        with open(target_path, "w", encoding="utf-8-sig", newline="") as out_f:
            writer = csv.writer(out_f, delimiter=";")
            writer.writerow(header11)
            writer.writerows(clean_rows)

        report = generator.validate_csv_detailed(
            str(target_path), strict=False, strict_overlap=False
        )
        if report.is_valid and report.register_count >= 3:
            generated_files.add(target_name.lower())
            added_local += 1
            total_tags_assigned += sum(1 for r in clean_rows if r[6])
        else:
            if target_path.exists():
                target_path.unlink()
    except Exception as e:
        print(f"  Error on local {l_csv.name}: {e}")

print(f"Local definitions tagged & saved: {added_local}")

# -------------------------------------------------------------------------
# Phase 2: Direct Webdyn definition CSVs from Equipementiers share
# -------------------------------------------------------------------------
print("\n[Phase 2] Mining and tagging definitions from Equipementiers share...")
share_csvs = [
    p for p in EQUIP_DIR.rglob("*.csv") if not p.name.startswith("~$")
]
added_share_csv = 0

for s_csv in share_csvs:
    try:
        rel_parts = s_csv.relative_to(EQUIP_DIR).parts
        mfg = rel_parts[0] if len(rel_parts) > 1 else "Generic"

        # Strictly open read-only
        with open(s_csv, "r", encoding="utf-8-sig", errors="ignore") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader, None)
            if not header or len(header) < 2 or not any(header):
                continue
            first_row = next(reader, None)
            if not first_row or len(first_row) < 7:
                continue
            fc_val = first_row[1].strip()
            if fc_val not in ("1", "2", "3", "4", "01", "02", "03", "04"):
                continue
            rows = [first_row] + list(reader)

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
        clean_rows = build_and_tag_rows(rows, detected_type=raw_cat or "Inverter")
        if len(clean_rows) < 3:
            continue

        header11 = (header + [""] * 11)[:11]

        with open(target_path, "w", encoding="utf-8-sig", newline="") as out_f:
            writer = csv.writer(out_f, delimiter=";")
            writer.writerow(header11)
            writer.writerows(clean_rows)

        report = generator.validate_csv_detailed(
            str(target_path), strict=False, strict_overlap=False
        )
        if report.is_valid and report.register_count >= 3:
            generated_files.add(target_filename.lower())
            added_share_csv += 1
            total_tags_assigned += sum(1 for r in clean_rows if r[6])
        else:
            if target_path.exists():
                target_path.unlink()

    except Exception:
        continue

print(f"Share CSV definitions tagged & saved: {added_share_csv}")

# -------------------------------------------------------------------------
# Phase 3: Excel Modbus tables from Equipementiers share (.xlsx)
# -------------------------------------------------------------------------
print("\n[Phase 3] Extracting, compiling & tagging Excel Modbus tables (.xlsx)...")
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

        # Format rows for builder
        raw_rows = []
        for r in mapped:
            raw_rows.append(
                [
                    "",
                    "3"
                    if str(r.get("RegisterType", "Holding")).lower() == "holding"
                    else "4",
                    str(r.get("Address", "")),
                    str(r.get("Format", "U16")),
                    "",
                    str(r.get("Name", "")),
                    str(r.get("Tag", "")),
                    str(r.get("ScaleFactor", "1.0")),
                    "0.0",
                    str(r.get("Unit", "")),
                    str(r.get("Action", "4")),
                ]
            )

        category = (
            "Inverter"
            if pfx == "INV"
            else ("Meter" if pfx == "MET" else ("Sensor" if pfx == "SEN" else "Other"))
        )
        clean_rows = build_and_tag_rows(raw_rows, detected_type=category)
        if len(clean_rows) < 3:
            continue

        target_path = OUTPUT_DIR / target_filename
        header11 = [
            "modbusRTU",
            category,
            mfg,
            base_name[:30],
            "",
            "",
            "",
            "",
            "",
            "",
            "",
        ]

        with open(target_path, "w", encoding="utf-8-sig", newline="") as out_f:
            writer = csv.writer(out_f, delimiter=";")
            writer.writerow(header11)
            writer.writerows(clean_rows)

        report = generator.validate_csv_detailed(
            str(target_path), strict=False, strict_overlap=False
        )
        if report.is_valid and report.register_count >= 3:
            generated_files.add(target_filename.lower())
            added_xlsx += 1
            total_tags_assigned += sum(1 for r in clean_rows if r[6])
        else:
            if target_path.exists():
                target_path.unlink()

    except Exception:
        continue

print(f"Excel Modbus definitions tagged & saved: {added_xlsx}")

# -------------------------------------------------------------------------
# Phase 4: Vendor CSV tables
# -------------------------------------------------------------------------
print("\n[Phase 4] Converting & tagging vendor CSV tables...")
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

        raw_rows = []
        for r in mapped:
            raw_rows.append(
                [
                    "",
                    "3"
                    if str(r.get("RegisterType", "Holding")).lower() == "holding"
                    else "4",
                    str(r.get("Address", "")),
                    str(r.get("Format", "U16")),
                    "",
                    str(r.get("Name", "")),
                    str(r.get("Tag", "")),
                    str(r.get("ScaleFactor", "1.0")),
                    "0.0",
                    str(r.get("Unit", "")),
                    str(r.get("Action", "4")),
                ]
            )

        category = (
            "Inverter"
            if pfx == "INV"
            else ("Meter" if pfx == "MET" else ("Sensor" if pfx == "SEN" else "Other"))
        )
        clean_rows = build_and_tag_rows(raw_rows, detected_type=category)
        if len(clean_rows) < 3:
            continue

        target_path = OUTPUT_DIR / target_filename
        header11 = [
            "modbusRTU",
            category,
            mfg,
            base_name[:30],
            "",
            "",
            "",
            "",
            "",
            "",
            "",
        ]

        with open(target_path, "w", encoding="utf-8-sig", newline="") as out_f:
            writer = csv.writer(out_f, delimiter=";")
            writer.writerow(header11)
            writer.writerows(clean_rows)

        report = generator.validate_csv_detailed(
            str(target_path), strict=False, strict_overlap=False
        )
        if report.is_valid and report.register_count >= 3:
            generated_files.add(target_filename.lower())
            added_vendor_csv += 1
            total_tags_assigned += sum(1 for r in clean_rows if r[6])
        else:
            if target_path.exists():
                target_path.unlink()

    except Exception:
        continue

print(f"Vendor CSV tables tagged & saved: {added_vendor_csv}")

# -------------------------------------------------------------------------
# Phase 5: Integrating and enriching definitions from generated_definitions/
# -------------------------------------------------------------------------
print("\n[Phase 5] Integrating and tagging definitions from generated_definitions/ ...")
gen_defs_dir = WORKSPACE_DIR / "generated_definitions"
added_gen = 0
for g_file in sorted(gen_defs_dir.glob("*.csv")):
    target_name = g_file.name
    if target_name.lower() in generated_files:
        continue
    target_path = OUTPUT_DIR / target_name
    try:
        with open(g_file, "r", encoding="utf-8-sig", errors="ignore") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader, None)
            if not header or len(header) < 2:
                continue
            rows = list(reader)

        eq_type = header[1].strip() if len(header) > 1 and header[1].strip() else "Inverter"
        clean_rows = build_and_tag_rows(rows, detected_type=eq_type)
        if len(clean_rows) < 3:
            continue

        header11 = (header + [""] * 11)[:11]

        with open(target_path, "w", encoding="utf-8-sig", newline="") as out_f:
            writer = csv.writer(out_f, delimiter=";")
            writer.writerow(header11)
            writer.writerows(clean_rows)

        report = generator.validate_csv_detailed(
            str(target_path), strict=False, strict_overlap=False
        )
        if report.is_valid and report.register_count >= 3:
            generated_files.add(target_name.lower())
            added_gen += 1
            total_tags_assigned += sum(1 for r in clean_rows if r[6])
        else:
            if target_path.exists():
                target_path.unlink()
    except Exception as e:
        print(f"Error on {g_file.name}: {e}")

print(f"Definitions from generated_definitions/ tagged & saved: {added_gen}")

# -------------------------------------------------------------------------
# Phase 6: Fox ESS Battery & Meter parts
# -------------------------------------------------------------------------
print("\n[Phase 6] Ensuring Fox ESS Battery and Meter parts are integrated...")
fox_files = [
    ("BAT_FoxESS_P3_H3_PRO.csv", "Battery"),
    ("MET_FoxESS_P3_H3_PRO.csv", "Meter"),
    ("BAT_FoxESS_RSeries.csv", "Battery"),
]
for fname, cat in fox_files:
    if fname.lower() in generated_files:
        continue
    src_f = gen_defs_dir / fname
    if src_f.exists():
        target_f = OUTPUT_DIR / fname
        try:
            with open(src_f, "r", encoding="utf-8-sig", errors="ignore") as f:
                reader = csv.reader(f, delimiter=";")
                header = next(reader)
                rows = list(reader)

            clean_rows = build_and_tag_rows(rows, detected_type=cat)
            header11 = (header + [""] * 11)[:11]

            with open(target_f, "w", encoding="utf-8-sig", newline="") as out_f:
                writer = csv.writer(out_f, delimiter=";")
                writer.writerow(header11)
                writer.writerows(clean_rows)

            report = generator.validate_csv_detailed(
                str(target_f), strict=False, strict_overlap=False
            )
            if report.is_valid:
                generated_files.add(fname.lower())
                total_tags_assigned += sum(1 for r in clean_rows if r[6])
        except Exception as e:
            print(f"Error on Fox file {fname}: {e}")

# -------------------------------------------------------------------------
# Final Summary & Statistics
# -------------------------------------------------------------------------
from collections import Counter

all_final = sorted(OUTPUT_DIR.glob("*.csv"))
prefix_counts = Counter()
total_registers = 0
total_sunpm_tags = 0
canonical_tags_distribution = Counter()

invalid_count = 0
for fpath in all_final:
    pfx = fpath.name.split("_")[0] if "_" in fpath.name else "OTH"
    prefix_counts[pfx] += 1
    report = generator.validate_csv_detailed(str(fpath), strict=False, strict_overlap=False)
    if not report.is_valid:
        invalid_count += 1
    with open(fpath, "r", encoding="utf-8-sig", errors="ignore") as fp:
        rdr = csv.reader(fp, delimiter=";")
        next(rdr, None)
        for r in rdr:
            if r and len(r) > 6:
                total_registers += 1
                tag = r[6].strip()
                if tag:
                    total_sunpm_tags += 1
                    canonical_tags_distribution[tag] += 1

print("\n" + "=" * 80)
print(f"TOTAL DEFINITIONS GENERATED IN '{OUTPUT_DIR.name}/': {len(all_final):,}")
print(f"TOTAL MODBUS REGISTERS: {total_registers:,}")
print(f"TOTAL CANONICAL SUNPM TAGS ASSIGNED: {total_sunpm_tags:,} (Tag Density: {total_sunpm_tags/max(1, total_registers):.1%})")
print(f"INVALID FILES COUNT: {invalid_count} (100% Valid Webdyn Schema: {invalid_count == 0})")
print("-" * 80)
print("Breakdown by Equipment Prefix:")
for pfx, cnt in sorted(prefix_counts.items()):
    print(f"  - {pfx:4s}: {cnt:4d} files")
print("-" * 80)
print("Top 20 Canonical SunPM Tags Assigned:")
for tag, cnt in canonical_tags_distribution.most_common(20):
    print(f"  - {tag:25s}: {cnt:5d}")
print("=" * 80)

# -------------------------------------------------------------------------
# Phase 7: Automated Technical Report Generation
# -------------------------------------------------------------------------
print("\n[Phase 7] Generating comprehensive technical reports...")
from DefFileGenerator.report_generator import (
    analyze_definitions_directory,
    export_report_json,
    generate_directory_technical_report,
)

report_md_dest = OUTPUT_DIR / "RAPPORT_TECHNIQUE_GENERATION.md"
docs_report_dest = WORKSPACE_DIR / "docs" / "RAPPORT_TECHNIQUE_GENERATION_DEFINITIONS.md"
report_json_dest = OUTPUT_DIR / "generation_report.json"

generate_directory_technical_report(OUTPUT_DIR, output_report_path=report_md_dest)
generate_directory_technical_report(OUTPUT_DIR, output_report_path=docs_report_dest)

dir_stats = analyze_definitions_directory(OUTPUT_DIR)
export_report_json(dir_stats, report_json_dest)

print(f"Technical report written to: {report_md_dest}")
print(f"Technical report mirrored to docs/: {docs_report_dest}")
print(f"JSON metrics written to: {report_json_dest}")
print("\nGeneration and Technical Reporting Completed Successfully.")

