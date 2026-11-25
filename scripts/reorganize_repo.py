#!/usr/bin/env python3
"""
Repository-Wide Architecture Organizer & WebdynSun Dual-Format Builder.

Enforces the canonical Webdyn definition taxonomy across the entire repository:
    <root>/<Manufacturer>/<Range>/
        <Model>.csv     (WebdynSunPM 11-column CSV, 5-Pillar compliant)
        <Model>.ini     (WebdynSun Classic .ini, MI v2.0 compliant)
        sources/        (Matching manufacturer manuals/tables + README.md provenance)

Also preserves authoritative .ini files from \\WDN-FILES.WEBDYN.COM (read-only)
and populates range source folders with matching PDF/XLSX/XML/JSON/TXT manuals.
"""

from __future__ import annotations

import csv
import hashlib
import os
import re
import shutil
import stat
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from DefFileGenerator.webdynsun_ini import (  # noqa: E402
    convert_sunpm_csv_to_webdynsun_ini,
    verify_webdynsun_ini,
)

LOCAL_SOURCES_ROOT = REPO_ROOT / "sources"
SHARE_ROOT = Path(r"\\WDN-FILES.WEBDYN.COM\Support\Equipementiers")
MAX_SOURCE_COPY_BYTES = 20 * 1024 * 1024  # 20 MiB cap per individual manual copy
VALID_SOURCE_EXTS = frozenset({".pdf", ".xlsx", ".xls", ".xlsm", ".xml", ".json", ".txt", ".csv"})

MANUFACTURER_ALIASES: dict[str, str] = {
    "foxess": "FoxESS",
    "fox_ess": "FoxESS",
    "fox ess": "FoxESS",
    "goodwe": "GoodWe",
    "carlogavazzi": "CarloGavazzi",
    "carlo_gavazzi": "CarloGavazzi",
    "carlo gavazzi": "CarloGavazzi",
    "marsenergy": "MarsEnergy",
    "mars_energy": "MarsEnergy",
    "marsrenewable": "MarsEnergy",
    "schneider": "Schneider",
    "schneider_electric": "Schneider",
    "sma": "SMA",
    "abb": "ABB",
    "huawei": "Huawei",
    "solplanet": "Solplanet",
    "sungrow": "Sungrow",
    "growatt": "Growatt",
    "fronius": "Fronius",
    "solaredge": "SolarEdge",
    "socomec": "Socomec",
    "chint": "Chint",
    "chint power": "Chint",
    "dyness": "Dyness",
    "x7": "X7",
    "kaco": "Kaco",
    "ingeteam": "Ingeteam",
    "delta": "Delta",
    "solis": "Solis",
    "ginlong_solis": "Solis",
    "saj": "SAJ",
    "janitza": "Janitza",
    "legrand": "Legrand",
    "cyberpower": "CyberPower",
    "acrel": "Acrel",
    "victron": "Victron",
    "deye": "Deye",
    "deye inverter": "Deye",
    "kehua": "Kehua",
    "sofarsolar": "SofarSolar",
    "solax": "SolaX",
    "kostal": "Kostal",
    "webdyn": "Webdyn",
    "alrosolar": "AlroSolar",
    "alro solar": "AlroSolar",
    "hoymiles": "Hoymiles",
    "hyxi": "Hyxi",
    "hyxipower": "Hyxi",
}

SHARE_FOXESS_SOURCES: dict[str, list[Path]] = {
    "P3_H3_Pro": [
        SHARE_ROOT / "Fox ESS" / "Sources" / "FoxESS Modbus Protocol--20260701 (V1.05.05.00).pdf",
        SHARE_ROOT / "Fox ESS" / "Sources" / "FoxESS-P3-5.0-SH-P3-15.0-SH-Oswiadczenie.pdf",
        SHARE_ROOT / "Fox ESS" / "RAPPORT_TECHNIQUE_FOXESS_P3_H3_PRO.md",
    ],
    "P3_Pro": [
        SHARE_ROOT / "Fox ESS" / "Sources" / "FoxESS Modbus Protocol--20260701 (V1.05.05.00).pdf",
        SHARE_ROOT / "Fox ESS" / "Sources" / "FoxESS-P3-5.0-SH-P3-15.0-SH-Oswiadczenie.pdf",
    ],
    "H3_Pro": [
        SHARE_ROOT / "Fox ESS" / "Sources" / "FoxESS Modbus Protocol--20260701 (V1.05.05.00).pdf",
    ],
    "G_MAX": [
        SHARE_ROOT / "Fox ESS" / "Sources" / "G-MAX Communication Protocol modbusTcp-EN-20250425(1).xlsx",
    ],
    "R_Series": [
        SHARE_ROOT / "Fox ESS" / "Sources" / "FoxESS RSeries xml file.xml",
        SHARE_ROOT / "Fox ESS" / "Sources" / "FR-R-Series-Datasheet-4.20.pdf",
        SHARE_ROOT / "Fox ESS" / "Sources" / "RSeries - Sunspec_Modbus_FOXESS_V1.04.pdf",
        SHARE_ROOT / "Fox ESS" / "Sources" / "RSeries - Sunspec_Modbus_FOXESS_V1.05.pdf",
        SHARE_ROOT / "Fox ESS" / "Sources" / "RSeries - Sunspec_Modbus_FOXESS_V1.06.pdf",
    ],
    "R_V_H3plus": [
        SHARE_ROOT / "Fox ESS" / "Sources" / "RVH3plus- Sunspec_Modbus_FOXESS_V1.07(20250909).pdf",
        SHARE_ROOT / "Fox ESS" / "Sources" / "R&V&H3plus&GmaxPlus-_Sunspec_Modbus_FOXESS_V1.08(20260514).pdf",
    ],
}


def _safe_dir_token(raw: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", raw.strip()).strip("_.")
    return cleaned or "General"


def _canonical_manufacturer(raw_mfg: str, csv_header_mfg: str = "") -> str:
    for candidate in (raw_mfg, csv_header_mfg):
        c = candidate.strip()
        if not c or c.lower() in {"unknown", "generic", "manufacturer", "vendor", "5", "k", "map", "power"}:
            continue
        key = c.lower().replace("-", "_")
        if key in MANUFACTURER_ALIASES:
            return MANUFACTURER_ALIASES[key]
        if c.lower() in MANUFACTURER_ALIASES:
            return MANUFACTURER_ALIASES[c.lower()]
        return _safe_dir_token(c)
    return "Generic"


def _canonical_foxess_range(name: str) -> str:
    u = name.upper()
    if "P3_H3_PRO" in u or "P3_H3PRO" in u or "P3H3" in u:
        return "P3_H3_Pro"
    if "H3_PRO" in u or "H3PRO" in u:
        return "H3_Pro"
    if "P3_PRO" in u or "P3PRO" in u:
        return "P3_Pro"
    if "R_V_H3PLUS" in u or "RV_H3PLUS" in u or "H3PLUS" in u:
        return "R_V_H3plus"
    if "G-MAX" in u or "G_MAX" in u or "GMAX" in u:
        return "G_MAX"
    if "RSERIES" in u or "R_SERIES" in u or "MODBUS_PROTOCOL" in u:
        return "R_Series"
    return "General"


def _canonical_range(mfg: str, filename: str, model_hdr: str = "") -> str:
    stem = Path(filename).stem
    if mfg == "FoxESS":
        return _canonical_foxess_range(stem)

    # Remove <PREFIX>_<Manufacturer>_ if present
    parts = stem.split("_")
    if len(parts) >= 3 and parts[0].upper() in {"INV", "MET", "SEN", "BAT", "COM", "OTH", "WPM"}:
        rest = "_".join(parts[2:])
    elif len(parts) >= 2 and parts[0].upper() in {"FOXESS", "FOX", "DYNESS", "X7", "MARS"}:
        rest = "_".join(parts[1:])
    else:
        rest = stem

    u = rest.upper()
    # Well-known industrial range groupings
    range_rules: tuple[tuple[str, str], ...] = (
        ("SUN2000", "SUN2000"),
        ("LUNA2000", "LUNA2000"),
        ("SMARTLOGGER", "SmartLogger"),
        ("SUNNYTRIPOWER", "Sunny_Tripower"),
        ("SUNNY_TRIPOWER", "Sunny_Tripower"),
        ("STP", "Sunny_Tripower"),
        ("CORE1", "Core1_Core2"),
        ("CORE2", "Core1_Core2"),
        ("SUNNYBOY", "Sunny_Boy"),
        ("SUNNY_BOY", "Sunny_Boy"),
        ("PVS800", "PVS800"),
        ("PVS-100", "PVS100_120"),
        ("PVS100", "PVS100_120"),
        ("TRIO", "TRIO"),
        ("TERRA", "Terra"),
        ("B21", "B21_B23_B24"),
        ("M4M", "M4M"),
        ("DTSU", "DTSU666"),
        ("DDSU666", "DDSU666"),
        ("NM8N", "NM8N_MCCB"),
        ("EM530", "EM530_EM540"),
        ("EM540", "EM530_EM540"),
        ("EM24", "EM24_EM300"),
        ("EM300", "EM24_EM300"),
        ("EM210", "EM210"),
        ("EM111", "EM111_EM112"),
        ("EM271", "EM271"),
        ("PM5000", "PM5000_IEM3000"),
        ("IEM3000", "PM5000_IEM3000"),
        ("DIRIS", "Diris_Countis"),
        ("COUNTIS", "Diris_Countis"),
        ("SYMO", "Symo_Primo_Gen24"),
        ("GEN24", "Symo_Primo_Gen24"),
        ("ASW", "ASW_Series"),
        ("AILOGGER", "Ai_Logger"),
        ("AI-LOGGER", "Ai_Logger"),
        ("SG_SERIES", "SG_Series"),
        ("SMT_SERIES", "SMT_Series"),
        ("ET_SERIES", "ET_Series"),
        ("LYNX", "Lynx_Home"),
        ("MAX_SERIES", "MAX_Series"),
        ("SE_SERIES", "SE_Series"),
        ("LOCALCONTROLLER", "LocalController"),
        ("LOCAL_CONTROLER", "LocalController"),
        ("LOCALCONTROLER", "LocalController"),
        ("DH100", "DH_Series"),
        ("DH200", "DH_Series"),
        ("DH_SERIES", "DH_Series"),
        ("X7", "X7_BESS"),
    )
    for needle, target_range in range_rules:
        if needle in u:
            return target_range

    # Fallback: strip generic role suffixes and take concise series name
    clean = re.sub(
        r"(_?(INVERTER|METER|BATTERY|BMS|EMS|STORAGE|ALL_IN_ONE|ALLINONE|MODBUS|RTU|TCP|SUNSPEC|WEBP|V\d+))$",
        "",
        rest,
        flags=re.IGNORECASE,
    )
    tokens = [t for t in clean.split("_") if t]
    if not tokens:
        return _safe_dir_token(model_hdr or "General")
    return _safe_dir_token("_".join(tokens[:2]))


def _read_csv_header(csv_path: Path) -> list[str]:
    try:
        with open(csv_path, encoding="utf-8-sig", errors="replace") as fp:
            return next(csv.reader(fp, delimiter=";"), [])
    except OSError:
        return []


def _sha256_short(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fp:
        for chunk in iter(lambda: fp.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _index_local_sources() -> dict[str, list[Path]]:
    """Index local source manuals under sources/ by lowercase manufacturer key."""
    idx: dict[str, list[Path]] = {}
    if not LOCAL_SOURCES_ROOT.exists():
        return idx
    for p in sorted(LOCAL_SOURCES_ROOT.rglob("*")):
        if not p.is_file() or p.suffix.lower() not in VALID_SOURCE_EXTS:
            continue
        if p.name in {"MANIFEST.json", "README.md"}:
            continue
        rel = p.relative_to(LOCAL_SOURCES_ROOT)
        parts = rel.parts
        mfg_tok = parts[1] if len(parts) >= 2 and parts[0].lower() == "manuals" else parts[0]
        key = _canonical_manufacturer(mfg_tok).lower()
        idx.setdefault(key, []).append(p)
    return idx


_GENERIC_SOURCE_TOKENS = frozenset(
    {
        "general",
        "generic",
        "modbus",
        "protocol",
        "manual",
        "user",
        "guide",
        "series",
        "inverter",
        "meter",
        "battery",
        "sensor",
        "other",
        "logger",
        "communication",
        "comm",
        "rtu",
        "tcp",
        "sunspec",
        "data",
        "table",
        "map",
        "power",
        "solar",
        "energy",
        "storage",
        "bess",
        "pro",
        "all",
        "en",
        "fr",
        "v1",
        "v2",
        "v3",
        "v4",
    }
)


def _find_matching_sources(
    mfg: str,
    rng: str,
    csv_names: list[str],
    local_index: dict[str, list[Path]],
    share_checked: dict[str, list[Path]],
) -> list[Path]:
    """Find ONLY source documents that genuinely match the exact manufacturer + range/models."""
    matches: list[Path] = []
    if mfg == "FoxESS" and rng in SHARE_FOXESS_SOURCES:
        for sp in SHARE_FOXESS_SOURCES[rng]:
            if sp.is_file():
                matches.append(sp)

    # Build strict, non-generic tokens from range name and actual CSV model names
    raw_tokens = re.split(r"[_\-.]+", rng)
    for cname in csv_names:
        stem = Path(cname).stem
        parts = stem.split("_", 2)
        if len(parts) == 3:
            raw_tokens.extend(re.split(r"[_\-.]+", parts[2]))
    specific_tokens = {
        t.lower()
        for t in raw_tokens
        if len(t) >= 3 and t.lower() not in _GENERIC_SOURCE_TOKENS and t.lower() != mfg.lower()
    }

    def _doc_matches(doc_path: Path) -> bool:
        if not specific_tokens:
            return False
        doc_text = re.sub(r"[^a-z0-9]+", " ", str(doc_path).lower())
        doc_compact = re.sub(r"[^a-z0-9]+", "", str(doc_path).lower())
        for tok in specific_tokens:
            tok_compact = re.sub(r"[^a-z0-9]+", "", tok)
            if len(tok_compact) < 3:
                continue
            if f" {tok} " in f" {doc_text} " or (len(tok_compact) >= 4 and tok_compact in doc_compact):
                return True
        return False

    # Match local sources/ files strictly by range/model tokens
    for cf in local_index.get(mfg.lower(), []):
        if _doc_matches(cf):
            matches.append(cf)

    # Cached share scan per manufacturer (strictly filtered by range/model tokens)
    mfg_key = mfg.lower()
    if mfg_key not in share_checked:
        found: list[Path] = []
        if SHARE_ROOT.exists():
            for d in SHARE_ROOT.iterdir():
                if d.is_dir() and _canonical_manufacturer(d.name).lower() == mfg_key:
                    for f in sorted(d.rglob("*")):
                        if (
                            f.is_file()
                            and f.suffix.lower() in VALID_SOURCE_EXTS
                            and f.name.lower() != "thumbs.db"
                            and f.stat().st_size <= MAX_SOURCE_COPY_BYTES
                            and ("sources" in str(f).lower() or "manual" in str(f).lower())
                        ):
                            found.append(f)
                    break
        share_checked[mfg_key] = found

    for sf in share_checked[mfg_key]:
        if _doc_matches(sf):
            matches.append(sf)

    # Deduplicate by filename (prefer SHARE_FOXESS_SOURCES / local sources first)
    dedup: dict[str, Path] = {}
    for m in matches:
        if m.name not in dedup and m.is_file() and m.stat().st_size <= MAX_SOURCE_COPY_BYTES:
            dedup[m.name] = m
    return list(dedup.values())


def _write_sources_readme(
    sources_dir: Path,
    mfg: str,
    rng: str,
    copied_sources: list[Path],
    csv_files: list[Path],
    ini_files: list[Path],
) -> None:
    lines = [
        f"# {mfg} - {rng} Sources",
        "",
        f"Source documentation and traceability manifest for **{mfg} / {rng}**.",
        "",
        "## Generated Definition Files",
        "",
        "- **WebdynSunPM (`.csv`)**: " + ", ".join(f"`{f.name}`" for f in sorted(csv_files)),
        "- **WebdynSun (`.ini`)**: " + ", ".join(f"`{f.name}`" for f in sorted(ini_files)),
        "",
        "## Source Documents",
        "",
    ]
    if copied_sources:
        lines.append("| File | Size (bytes) | SHA-256 |")
        lines.append("|---|---:|---|")
        for sf in sorted(copied_sources):
            lines.append(f"| `{sf.name}` | {sf.stat().st_size} | `{_sha256_short(sf)}` |")
    else:
        lines.extend(
            [
                "Derived from the validated Webdyn manufacturer definition baseline / SunSpec model profile.",
                "No separate standalone PDF/XLSX manual file was bundled in local `sources/` for this legacy model.",
            ]
        )
    sources_dir.mkdir(parents=True, exist_ok=True)
    (sources_dir / "README.md").write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")


def organize_directory(
    root_dir: Path,
    local_index: dict[str, list[Path]],
    share_checked: dict[str, list[Path]],
    *,
    copy_binary_sources: bool = True,
) -> dict[str, int]:
    """Reorganize a definition root directory into <Manufacturer>/<Range>/{csv, ini, sources}."""
    stats = {"csv_moved": 0, "ini_generated": 0, "ini_invalid": 0, "sources_copied": 0, "ranges": 0}
    if not root_dir.exists():
        return stats

    all_csvs = sorted(
        p for p in root_dir.rglob("*.csv") if "sources" not in p.parts
    )
    range_buckets: dict[tuple[str, str], list[Path]] = {}

    for csv_path in all_csvs:
        hdr = _read_csv_header(csv_path)
        hdr_mfg = hdr[2] if len(hdr) > 2 else ""
        hdr_model = hdr[3] if len(hdr) > 3 else ""

        parts = csv_path.stem.split("_")
        raw_mfg = parts[1] if len(parts) > 1 else ""
        mfg = _canonical_manufacturer(raw_mfg, hdr_mfg)
        rng = _canonical_range(mfg, csv_path.name, hdr_model)

        dest_dir = root_dir / mfg / rng
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest_csv = dest_dir / csv_path.name

        if csv_path.resolve() != dest_csv.resolve():
            try:
                if dest_csv.exists():
                    # Keep existing if identical, else replace cleanly
                    csv_path.unlink()
                else:
                    shutil.move(str(csv_path), str(dest_csv))
                stats["csv_moved"] += 1
            except OSError:
                # Copy fallback if source handle is locked by an editor
                if not dest_csv.exists():
                    shutil.copy2(str(csv_path), str(dest_csv))

        range_buckets.setdefault((mfg, rng), []).append(dest_csv)

    # Remove old empty subdirectories left behind inside root_dir
    for d in sorted([p for p in root_dir.rglob("*") if p.is_dir()], reverse=True):
        if "sources" in d.parts:
            continue
        try:
            if not any(d.iterdir()):
                d.rmdir()
        except OSError:
            pass

    # For each <Manufacturer>/<Range> directory: generate WebdynSun .ini + populate sources/
    for (mfg, rng), csv_list in sorted(range_buckets.items()):
        stats["ranges"] += 1
        range_dir = root_dir / mfg / rng
        ini_list: list[Path] = []
        for csv_file in csv_list:
            ini_file = csv_file.with_suffix(".ini")
            convert_sunpm_csv_to_webdynsun_ini(csv_file, ini_file)
            ok, _issues = verify_webdynsun_ini(ini_file)
            if ok:
                stats["ini_generated"] += 1
            else:
                stats["ini_invalid"] += 1
            ini_list.append(ini_file)

        # Preserve official FoxESS share .ini reference files in P3_H3_Pro / R_Series
        if mfg == "FoxESS":
            ref_map = {
                "P3_H3_Pro": SHARE_ROOT / "Fox ESS" / "P3-H3-PRO" / "FoxESS_P3_H3_PRO_INVERTER.ini",
                "R_Series": SHARE_ROOT / "Fox ESS" / "R SERIES" / "INV_FOXESS_RSERIESINVERTER.ini",
            }
            if rng in ref_map and ref_map[rng].is_file():
                ref_dest = range_dir / f"REF_{ref_map[rng].name}"
                if not ref_dest.exists():
                    shutil.copy2(str(ref_map[rng]), str(ref_dest))

        sources_dir = range_dir / "sources"
        sources_dir.mkdir(parents=True, exist_ok=True)
        if copy_binary_sources:
            # Remove any previously misplaced non-README binary files before copying strict matches
            for old_f in list(sources_dir.iterdir()):
                if old_f.is_file() and old_f.name != "README.md":
                    os.chmod(old_f, stat.S_IWRITE)
                    old_f.unlink()
            for src_doc in _find_matching_sources(
                mfg, rng, [c.name for c in csv_list], local_index, share_checked
            ):
                dst_doc = sources_dir / src_doc.name
                shutil.copyfile(str(src_doc), str(dst_doc))
                stats["sources_copied"] += 1
        copied_sources: list[Path] = [
            f for f in sources_dir.iterdir() if f.is_file() and f.name != "README.md"
        ]

        _write_sources_readme(sources_dir, mfg, rng, copied_sources, csv_list, ini_list)

    return stats


def main() -> int:
    local_index = _index_local_sources()
    share_checked: dict[str, list[Path]] = {}

    targets = [
        (REPO_ROOT / "production", True),
        (REPO_ROOT / "Definitions", False),
        (REPO_ROOT / "generated_definitions", False),
        (REPO_ROOT / "tagged_definitions", False),
    ]
    for root_dir, copy_bins in targets:
        print(f"Organizing {root_dir.name}/ ...")
        st = organize_directory(
            root_dir, local_index, share_checked, copy_binary_sources=copy_bins
        )
        print(f"  -> {st}")

    # Organize dedicated bundle directories in-place (adding .ini + sources/)
    bundles = [
        REPO_ROOT / "FoxESS_P3_H3_Pro_Definitions",
        REPO_ROOT / "Dyness_DH_Series_Definitions",
        REPO_ROOT / "MarsRenewable_Definitions",
        REPO_ROOT / "X7_Definitions",
    ]
    for b in bundles:
        if not b.exists():
            continue
        csvs = sorted(p for p in b.rglob("*.csv") if "sources" not in p.parts)
        for cf in csvs:
            convert_sunpm_csv_to_webdynsun_ini(cf, cf.with_suffix(".ini"))
        # Add FoxESS sources to FoxESS_P3_H3_Pro_Definitions/*/sources
        if b.name == "FoxESS_P3_H3_Pro_Definitions":
            for sub in ("P3_Pro", "H3_Pro", "Commun", "Autres_Gammes_FoxESS"):
                sub_dir = b / sub
                if sub_dir.is_dir():
                    s_dir = sub_dir / "sources"
                    s_dir.mkdir(exist_ok=True)
                    src_pdf = LOCAL_SOURCES_ROOT / "manuals" / "FoxESS"
                    if src_pdf.is_dir():
                        for sf in src_pdf.iterdir():
                            if sf.is_file() and not (s_dir / sf.name).exists():
                                shutil.copy2(str(sf), str(s_dir / sf.name))
        print(f"Bundle updated with WebdynSun .ini + sources/: {b.name} ({len(csvs)} CSVs)")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
