#!/usr/bin/env python3
"""
Deep comparison benchmark between generated definitions and the R&D Bible in DEF_PM_R&D.
"""

import os
import sys
import glob
import re
import csv
import io
from pathlib import Path
from typing import Any, Dict, List, Tuple

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator

DEF_PM_DIR = r"G:\My Drive\08092026\DEF_PM_R&D"
EQUIP_DIR = r"G:\My Drive\08092026\Equipementiers"

def load_validated_def(filepath: str) -> Dict[str, Any]:
    """Parse a validated Webdyn definition CSV file."""
    with open(filepath, "r", encoding="latin-1") as f:
        lines = [line.strip() for line in f if line.strip()]

    if not lines:
        return {"header": [], "registers": {}, "raw_count": 0}

    header = lines[0].split(";")
    registers = {}
    for idx, line in enumerate(lines[1:]):
        parts = line.split(";")
        if len(parts) >= 11:
            addr = parts[2].strip()
            # reg: Index, RegType, Address, Type, Info4, Name, Tag, CoefA, CoefB, Unit, Action
            registers[addr] = {
                "index": parts[0].strip(),
                "reg_type": parts[1].strip(),
                "address": addr,
                "type": parts[3].strip(),
                "info4": parts[4].strip(),
                "name": parts[5].strip(),
                "tag": parts[6].strip(),
                "coef_a": parts[7].strip(),
                "coef_b": parts[8].strip(),
                "unit": parts[9].strip(),
                "action": parts[10].strip(),
            }
        elif len(parts) >= 3:
            addr = parts[2].strip()
            registers[addr] = {"address": addr, "raw": parts}

    return {
        "header": header,
        "registers": registers,
        "raw_count": len(registers),
    }

def find_candidate_sources_for_def(def_path: Path, equip_index: List[Path]) -> List[Path]:
    """Find source documents corresponding to a validated definition."""
    def_name = def_path.stem.lower()
    # Normalize name: strip prefixes like INV_, METER_, SEN_, COM_, OTH_, DISPLAY_, ANAEX_
    norm_name = re.sub(r"^(inv|meter|sen|sensor|com|oth|display|anaex|stringbox)_", "", def_name)
    parts = [p for p in re.split(r"[_\-\s]+", norm_name) if len(p) >= 3]

    matched = []
    for sp in equip_index:
        sp_lower = sp.name.lower()
        sp_dir = str(sp.parent).lower()
        # skip temporary, ticket, support
        if any(ign in sp_lower for ign in ["ticket", "support", "reunion", "prio", "qualité", "client", "trame"]):
            continue

        score = 0
        # If manufacturer is in parent directory
        for p in parts:
            if p in sp_dir:
                score += 2
            if p in sp_lower:
                score += 3

        if score >= 5:
            matched.append((score, sp))

    matched.sort(key=lambda x: x[0], reverse=True)
    return [p for _, p in matched]

def run_benchmark():
    print("=== Scanning DEF_PM_R&D ===")
    def_files = list(Path(DEF_PM_DIR).rglob("*.csv"))
    print(f"Total validated definitions found: {len(def_files)}")

    print("\n=== Indexing Technical Sources in Equipementiers ===")
    equip_sources = []
    for ext in ("*.xlsx", "*.xls", "*.pdf", "*.csv", "*.xml"):
        equip_sources.extend(list(Path(EQUIP_DIR).rglob(ext)))
    print(f"Total candidate technical files indexed: {len(equip_sources)}")

    extractor = Extractor()
    generator = Generator()

    results = []

    # Let's test high-value representative R&D definitions across all equipment categories
    for df in sorted(def_files):
        val_data = load_validated_def(str(df))
        val_reg_count = val_data["raw_count"]
        header = val_data["header"]
        category = header[1] if len(header) > 1 else df.parent.name
        mfg = header[2] if len(header) > 2 else ""
        model = header[3] if len(header) > 3 else df.stem

        candidates = find_candidate_sources_for_def(df, equip_sources)
        if not candidates:
            results.append({
                "def_file": df.name,
                "category": category,
                "mfg": mfg,
                "model": model,
                "val_count": val_reg_count,
                "gen_count": 0,
                "source_file": "Source non trouvée",
                "status": "NO_SOURCE",
                "common_addrs": 0,
                "match_rate": 0.0,
            })
            continue

        # Try candidates to see if we can extract registers
        best_gen_regs = {}
        best_candidate = None

        for cand in candidates[:5]:
            cand_ext = cand.suffix.lower()
            try:
                tables = []
                if cand_ext in (".xlsx", ".xls"):
                    tables = list(extractor.extract_from_excel(str(cand)))
                elif cand_ext == ".pdf":
                    tables = list(extractor.extract_from_pdf(str(cand)))
                elif cand_ext == ".csv":
                    tables = list(extractor.extract_from_csv(str(cand)))
                elif cand_ext == ".xml":
                    tables = list(extractor.extract_from_xml(str(cand)))

                if not tables:
                    continue

                mapped = extractor.map_and_clean(tables)
                gen_rows = list(generator.process_rows(mapped))
                if len(gen_rows) > len(best_gen_regs):
                    best_gen_regs = {r["Info2"].split("_")[0]: r for r in gen_rows if "Info2" in r}
                    best_candidate = cand
            except Exception:
                continue

        val_addrs = set(val_data["registers"].keys())
        gen_addrs = set(best_gen_regs.keys())
        common = val_addrs.intersection(gen_addrs)
        
        match_rate = (len(common) / len(val_addrs) * 100.0) if val_addrs else 0.0

        results.append({
            "def_file": df.name,
            "category": category,
            "mfg": mfg,
            "model": model,
            "val_count": val_reg_count,
            "gen_count": len(best_gen_regs),
            "source_file": best_candidate.name if best_candidate else "Aucune extraction valide",
            "source_rel": str(best_candidate.relative_to(Path(EQUIP_DIR))) if best_candidate else "",
            "status": "SUCCESS" if common else ("GEN_ONLY" if gen_addrs else "NO_EXTRACTION"),
            "common_addrs": len(common),
            "match_rate": match_rate,
            "sample_common": list(common)[:5],
        })

    # Summary
    success_matches = [r for r in results if r["status"] == "SUCCESS"]
    print(f"\n================ BENCHMARK SUMMARY ================")
    print(f"Total Validated Definitions: {len(results)}")
    print(f"Definitions directly matched & extracted from technical sources: {len(success_matches)}")
    print(f"Definitions with common Modbus addresses: {len(success_matches)}")

    for r in success_matches[:25]:
        print(f"\n[MATCH] {r['def_file']} ({r['mfg']} {r['model']})")
        print(f"  Source: {r['source_file']}")
        print(f"  R&D Count: {r['val_count']} | Gen Count: {r['gen_count']} | Common Addrs: {r['common_addrs']} ({r['match_rate']:.1f}%)")
        if r['sample_common']:
            print(f"  Sample Addrs: {r['sample_common']}")

if __name__ == "__main__":
    run_benchmark()
