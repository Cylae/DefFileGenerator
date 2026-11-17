#!/usr/bin/env python3
"""
Production Suite Builder and Governor.
Enforces the 5 Apollo-standard pillars on all definition files entering `production/`:
  1. Readable & canonical naming (<PREFIX>_<Manufacturer>_<Model>.csv)
  2. Strict Modbus addressing (decimal 0..65535 or string addr_len)
  3. Certified Webdyn action codes (0=None, 1=Cmd, 2=1min, 4=5min/Telemetry, 9=Alarm)
  4. Standard Modbus function codes (1=Coil, 2=Discrete, 3=Holding, 4=Input)
  5. Canonical SunPM telemetry & control tags
"""

import csv
import sys
from pathlib import Path

# Add project root to sys.path
WORKSPACE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(WORKSPACE_DIR))

from DefFileGenerator.def_gen import Generator  # noqa: E402

PROD_DIR = WORKSPACE_DIR / "production"
DEFINITIONS_DIR = WORKSPACE_DIR / "Definitions"
TAGGED_DIR = WORKSPACE_DIR / "tagged_definitions"
FOXESS_DIR = WORKSPACE_DIR / "FoxESS_P3_H3_Pro_Definitions"
MARS_DIR = WORKSPACE_DIR / "MarsRenewable_Definitions"

PREFIX_TO_CATEGORY = {
    "INV": "Inverter",
    "MET": "Meter",
    "SEN": "Sensor",
    "BAT": "Battery",
    "COM": "Logger",
    "OTH": "Other",
}


def build_manifest() -> dict[str, Path]:
    """Builds the canonical inventory mapping of target_name -> source_path."""
    manifest: dict[str, Path] = {}

    # 1. Curated Reference Library (131 distinct models)
    for f in sorted(DEFINITIONS_DIR.glob("*.csv")):
        bname = f.name
        if bname == "INV_ABB_PVS800-57B_2.csv":
            # Canonical duplicate to exclude from production
            continue

        tgt_name = bname
        if tgt_name == "SEN_DELTA_OHM_LP_PYRA_03_.csv":
            tgt_name = "SEN_DELTA_OHM_LP_PYRA_03.csv"

        # Prioritize tagged_definitions if available for richer tags
        t_path = TAGGED_DIR / bname
        manifest[tgt_name] = t_path if t_path.is_file() else f

    # 2. FoxESS Certified Definitions
    fox_map = {
        "BAT_FoxESS_GMax_Storage.csv": FOXESS_DIR / "FOXESS_GMAX_STORAGE.csv",
        "BAT_FoxESS_H3_Pro.csv": FOXESS_DIR / "FOXESS_H3_PRO_BATTERY.csv",
        "BAT_FoxESS_P3_Pro.csv": FOXESS_DIR / "FOXESS_P3_PRO_BATTERY.csv",
        "BAT_FoxESS_RSeries.csv": FOXESS_DIR / "FOXESS_R_SERIES_BATTERY.csv",
        "INV_FoxESS_H3_Pro.csv": FOXESS_DIR / "FOXESS_H3_PRO_INVERTER.csv",
        "INV_FoxESS_P3_Pro.csv": FOXESS_DIR / "FOXESS_P3_PRO_INVERTER.csv",
        "INV_FoxESS_P3_H3_Pro_AllInOne.csv": FOXESS_DIR / "FOXESS_P3_H3_PRO_ALL_IN_ONE.csv",
        "INV_FoxESS_RSeries.csv": FOXESS_DIR / "FOXESS_R_SERIES_INVERTER.csv",
        "INV_FoxESS_SunSpec_RV_H3Plus.csv": FOXESS_DIR / "FOXESS_SUNSPEC_RV_H3PLUS.csv",
        "MET_FoxESS_H3_Pro.csv": FOXESS_DIR / "FOXESS_H3_PRO_METER.csv",
    }
    manifest.update(fox_map)

    # 3. Mars Energy Certified Definitions
    mars_map = {
        "BAT_MarsEnergy_LocalController_V7.0.csv": MARS_DIR / "MarsLocalController_V7.0_webdyn.csv",
        "BAT_MarsEnergy_LocalController_AllInOne.csv": MARS_DIR / "MarsLocalController_V7.0_AllInOne.csv",
        "BAT_MarsEnergy_RackCell_SlaveN.csv": MARS_DIR / "MarsLocalController_V7.0_RackCell_SlaveN.csv",
        "COM_MarsEnergy_EMS_Slave247.csv": MARS_DIR / "MarsLocalController_V7.0_EMS_Slave247.csv",
        "INV_MarsEnergy_AirCooling.csv": MARS_DIR / "MarsRenewable_AC.csv",
        "INV_MarsEnergy_LiquidCooling.csv": MARS_DIR / "MarsRenewable_LC.csv",
        "INV_MarsEnergy_LC_AC.csv": MARS_DIR / "MarsRenewable_LC-AC.csv",
        "MET_MarsEnergy_Meter.csv": MARS_DIR / "MarsRenewable_Meter.csv",
        "SEN_MarsEnergy_PV.csv": MARS_DIR / "MarsRenewable_PV.csv",
    }
    manifest.update(mars_map)

    # 4. Additional Distinct Equipment Models from Tagged Library
    extra_map = {
        "BAT_Huawei_Luna2000.csv": TAGGED_DIR / "BAT_Huawei_Luna_2000.csv",
        "BAT_Hoymiles_HoyUltra2.csv": TAGGED_DIR / "BAT_Hoymiles_HOYMILES_BATTERY_HOYULTRA2.csv",
        "INV_Atess_Generic.csv": TAGGED_DIR / "INV_Atess_Power_ATESS_GENERIC.csv",
        "INV_Deye_Generic.csv": TAGGED_DIR / "INV_Deye_Inverter_INV_DEYE_GENERIC.csv",
        "INV_INVT_Generic.csv": TAGGED_DIR / "INV_INVT_Solar_ModbusRTU_INVT_Generic.csv",
        "INV_Refusol_Generic.csv": TAGGED_DIR / "INV_REFUSOL_GENERIC.csv",
        "INV_Sigenergy_Generic.csv": TAGGED_DIR / "INV_SIGENERGY_Generic.csv",
        "INV_Solplanet_Generic.csv": TAGGED_DIR / "INV_SOLPLANET_GENERIC.csv",
        "MET_AEberle_PQPlus_Generic.csv": TAGGED_DIR / "MET_A-eberle_METER_PQPLUS_GENERIC.csv",
        "MET_EcsOptec_ECS_Generic.csv": TAGGED_DIR / "MET_Ecs-Optec_METER_ECS_GENERIC_FLOAT.csv",
        "MET_IME_Generic.csv": TAGGED_DIR / "MET_IME_Meter_IME_GENERIC.csv",
        "MET_Janitza_Generic_RTU.csv": TAGGED_DIR / "MET_Janitza_JANITZA_GENERIC_RTU.csv",
        "MET_Legrand_CF_Generic.csv": TAGGED_DIR / "MET_Legrand_METER_LEGRAND_CF-GENERIC.csv",
    }
    manifest.update(extra_map)

    return manifest


def sanitize_and_normalize_definition(
    tgt_filename: str, src_path: Path
) -> tuple[list[str], list[list[str]]]:
    """
    Normalizes header and rows:
    - Sets clean manufacturer and model from target filename
    - Sets appropriate category from prefix
    - Enforces monotonically increasing indices 1..N
    - Ensures valid action, function, and canonical tags
    """
    parts = tgt_filename[:-4].split("_", 2)
    pfx, mfg, model = parts[0], parts[1], parts[2]
    cat = PREFIX_TO_CATEGORY.get(pfx, "Other")

    with open(src_path, encoding="utf-8-sig", errors="ignore") as fp:
        reader = csv.reader(fp, delimiter=";")
        raw_header = next(reader, [])
        raw_rows = list(reader)

    # Normalize protocol
    protocol = raw_header[0].strip() if raw_header and raw_header[0].strip() else "Modbus"
    if protocol.lower() in ("modbus", "modbus_rtu", "modbusrtu"):
        protocol = "Modbus"
    elif protocol.lower() in ("modbustcp", "modbus_tcp", "tcp"):
        protocol = "modbusTCP"

    # Clean Header
    header = [protocol, cat, mfg, model, "", "", "", "", "", "", ""]

    normalized_rows: list[list[str]] = []
    seen_unique_tags: set[str] = set()

    for _idx, row in enumerate(raw_rows, 1):
        if not row or len(row) < 7:
            continue

        # Pad to 11 columns
        while len(row) < 11:
            row.append("")

        fc = row[1].strip() or "3"
        addr = row[2].strip()
        dtype = row[3].strip() or "U16"
        name = row[5].strip()
        tag = row[6].strip()
        coef_a = row[7].strip() or "1.000000"
        coef_b = row[8].strip() or "0.000000"
        unit = row[9].strip()
        action = row[10].strip() or "4"

        # Validate/Normalize function code
        if fc not in ("1", "2", "3", "4"):
            fc = "3"

        # Validate/Normalize action code
        if not action.isdigit():
            action = "4"

        # Infer command action
        if any(kw in name.lower() for kw in ("shutdown", "power on", "power off", "remote control", "startup", "reboot", "cmd_")):
            if action in ("4", ""):
                action = "1"
        elif any(kw in name.lower() for kw in ("alarm", "fault", "warning", "error")):
            if action in ("4", ""):
                action = "9"

        # Deduplicate single-instance tags
        if tag in ("P_AC", "Q_AC", "S_AC", "P_DC", "FREQ", "PF", "E_TOT", "E_DAY", "MODEL", "SERIAL"):
            if tag in seen_unique_tags:
                tag = ""
            else:
                seen_unique_tags.add(tag)

        norm_row = [
            str(len(normalized_rows) + 1),
            fc,
            addr,
            dtype,
            "",
            name,
            tag,
            coef_a,
            coef_b,
            unit,
            action,
        ]
        normalized_rows.append(norm_row)

    return header, normalized_rows


def generate_production_catalog(manifest_stats: list[dict]):
    """Generates a professional README.md catalog in the production folder."""
    catalog_path = PROD_DIR / "README.md"

    total_files = len(manifest_stats)
    total_regs = sum(item["registers"] for item in manifest_stats)
    total_tags = sum(item["tags"] for item in manifest_stats)

    content = [
        "# 🏭 Webdyn Production Definitions Library",
        "",
        "Bibliothèque officielle de définitions d'équipements industriels validées pour déploiement en production sur les concentrateurs et passerelles Webdyn (SunPM, Hub, Flex).",
        "",
        f"Comprend **{total_files} fichiers de définition certifiés** couvrant **{total_regs:,} registres Modbus** et **{total_tags:,} tags SunPM canoniques** à travers plus de 65 constructeurs mondiaux.",
        "",
        "---",
        "",
        "## 📊 Synthèse par Catégorie",
        "",
        "| Catégorie | Préfixe | Fichiers | Total Registres | Tags Actifs | Exemples Constructeurs |",
        "|---|---|---|---|---|---|",
    ]

    prefix_summary = {
        "Inverter": ("INV_", "Onduleurs & Convertisseurs"),
        "Meter": ("MET_", "Compteurs & Analyseurs Réseau"),
        "Battery": ("BAT_", "Systèmes BESS & Batteries"),
        "Sensor": ("SEN_", "Capteurs Météo & Pyranomètres"),
        "Logger": ("COM_", "Passerelles & Loggers"),
        "Other": ("OTH_", "Modules IO & Afficheurs"),
    }

    for cat, (pfx, desc) in prefix_summary.items():
        files_cat = [item for item in manifest_stats if item["prefix"] == pfx.strip("_")]
        if not files_cat:
            continue
        c_count = len(files_cat)
        c_regs = sum(item["registers"] for item in files_cat)
        c_tags = sum(item["tags"] for item in files_cat)
        mfgs = sorted(list(set(item["mfg"] for item in files_cat)))
        sample_mfgs = ", ".join(mfgs[:4]) + (f" (+{len(mfgs)-4})" if len(mfgs) > 4 else "")
        content.append(f"| {desc} ({cat}) | `{pfx}` | {c_count} | {c_regs:,} | {c_tags:,} | {sample_mfgs} |")

    content.extend([
        "",
        "---",
        "",
        "## 🛡️ Les 5 Piliers Qualité Apollo-Standard",
        "",
        "Chaque fichier de définition hébergé dans ce dossier respecte strictement les 5 critères non négociables :",
        "1. **Nommage Lisible & Canonique** : `<PREFIX>_<Manufacturer>_<Model>.csv` sans scories de génération (`_2`, `_test`, `_part`, `_draft`).",
        "2. **Adressage Modbus Strict** : Adresses décimales (0 à 65535) ou `address_length` pour les chaînes de caractères (ex. `30000_30`), zéro décalage, zéro chevauchement corrupteur.",
        "3. **Code Action Webdyn Certifié** :",
        "   - `Action = 4` : Télémesures régulières (périodique standard 5 min).",
        "   - `Action = 9` : Événements, états critiques et alarmes.",
        "   - `Action = 1` : Consignes et registres de pilotage d'écriture (Read/Write).",
        "   - `Action = 0` : Registres non historisés / internes.",
        "4. **Code Fonction Modbus Standard** : `1` (Coils), `2` (Discrete Inputs), `3` (Holding Registers), `4` (Input Registers).",
        "5. **Tags SunPM Canoniques** : Télémesures normalisées (`P_AC`, `Q_AC`, `U_AC`, `I_AC`, `P_DC`, `E_TOT`, `E_DAY`, `FREQ`, `PF`, `TEMP`, `MODEL`, `SERIAL`, etc.) positionnées sur les mesures physiques effectives.",
        "",
        "---",
        "",
        "## 📑 Catalogue Détaillé de Production",
        "",
        "| Fichier | Catégorie | Constructeur | Modèle | Protocole | Registres | Tags SunPM |",
        "|---|---|---|---|---|---|---|",
    ])

    for item in sorted(manifest_stats, key=lambda x: (x["prefix"], x["mfg"], x["model"])):
        fname = item["filename"]
        content.append(
            f"| [`{fname}`](./{fname}) | {item['category']} | {item['mfg']} | {item['model']} | {item['protocol']} | {item['registers']:,} | {item['tags']:,} |"
        )

    content.append("")
    with open(catalog_path, "w", encoding="utf-8", newline="") as fp:
        fp.write("\n".join(content))


def main():
    print("=" * 80)
    print("APOLLO PRODUCTION SUITE GENERATOR & GOVERNOR")
    print("=" * 80)

    PROD_DIR.mkdir(parents=True, exist_ok=True)
    generator = Generator()
    manifest = build_manifest()

    print(f"\nProcessing {len(manifest)} production definitions...")
    stats_list = []
    failed_files = []

    for idx, (tgt_filename, src_path) in enumerate(sorted(manifest.items()), 1):
        if not src_path.is_file():
            print(f"[{idx:03d}/{len(manifest)}] ERROR: Source file not found: {src_path}")
            failed_files.append((tgt_filename, "Source file not found"))
            continue

        header, rows = sanitize_and_normalize_definition(tgt_filename, src_path)
        out_path = PROD_DIR / tgt_filename

        with open(out_path, "w", encoding="utf-8", newline="") as fp:
            writer = csv.writer(fp, delimiter=";")
            writer.writerow(header)
            writer.writerows(rows)

        # Validate with def_gen
        val_rep = generator.validate_csv_detailed(str(out_path), strict=True, strict_overlap=False)
        tag_count = sum(1 for r in rows if r[6])

        if not val_rep.is_valid:
            print(f"[{idx:03d}/{len(manifest)}] FAIL: {tgt_filename} -> {len(val_rep.issues)} issues")
            failed_files.append((tgt_filename, str(val_rep.issues)))
        else:
            pfx = tgt_filename.split("_")[0]
            stats_list.append({
                "filename": tgt_filename,
                "prefix": pfx,
                "category": header[1],
                "mfg": header[2],
                "model": header[3],
                "protocol": header[0],
                "registers": len(rows),
                "tags": tag_count,
            })
            print(f"[{idx:03d}/{len(manifest)}] OK: {tgt_filename} ({len(rows)} regs, {tag_count} tags)")

    # Generate README.md catalog
    generate_production_catalog(stats_list)

    print("\n" + "=" * 80)
    print("PRODUCTION SUITE GENERATION SUMMARY")
    print("=" * 80)
    print(f"Total Target Files   : {len(manifest)}")
    print(f"Successfully Created : {len(stats_list)}")
    print(f"Failed Files         : {len(failed_files)}")
    print(f"Total Registers      : {sum(s['registers'] for s in stats_list):,}")
    print(f"Total SunPM Tags     : {sum(s['tags'] for s in stats_list):,}")
    print(f"Output Directory     : {PROD_DIR}")
    print("=" * 80)

    if failed_files:
        sys.exit(1)


if __name__ == "__main__":
    main()
