#!/usr/bin/env python3
"""
Deep register-level comparison between generated definitions and the R&D Bible.
Produces exact numerical and semantic matching metrics.
"""

import os
import sys
import json
from pathlib import Path
from collections import defaultdict

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from DefFileGenerator.extractor import Extractor
from DefFileGenerator.def_gen import Generator
from scratch.compare_with_def_pm_rd import (
    load_validated_def,
    find_candidate_sources_for_def,
    DEF_PM_DIR,
    EQUIP_DIR,
)


def run_deep_comparison():
    def_files = sorted(list(Path(DEF_PM_DIR).rglob("*.csv")))
    equip_sources = []
    for ext in ("*.xlsx", "*.xls", "*.pdf", "*.csv", "*.xml"):
        equip_sources.extend(list(Path(EQUIP_DIR).rglob(ext)))

    extractor = Extractor()
    generator = Generator()

    category_stats = defaultdict(
        lambda: {
            "total_files": 0,
            "matched_files": 0,
            "rd_total_regs": 0,
            "gen_total_regs": 0,
            "common_regs": 0,
            "type_matches": 0,
            "coef_matches": 0,
            "name_matches": 0,
            "files_detail": [],
        }
    )

    equipment_matches = []

    for df in def_files:
        val_data = load_validated_def(str(df))
        cat = df.parent.name
        val_regs = val_data["registers"]
        val_count = len(val_regs)

        category_stats[cat]["total_files"] += 1
        category_stats[cat]["rd_total_regs"] += val_count

        candidates = find_candidate_sources_for_def(df, equip_sources)
        if not candidates:
            category_stats[cat]["files_detail"].append(
                {
                    "name": df.name,
                    "status": "NO_SOURCE",
                    "val_count": val_count,
                    "gen_count": 0,
                    "common": 0,
                    "match_pct": 0.0,
                }
            )
            continue

        best_gen_rows = []
        best_candidate = None
        for cand in candidates[:5]:
            ext = cand.suffix.lower()
            try:
                tables = []
                if ext in (".xlsx", ".xls"):
                    tables = list(extractor.extract_from_excel(str(cand)))
                elif ext == ".pdf":
                    tables = list(extractor.extract_from_pdf(str(cand)))
                elif ext == ".csv":
                    tables = list(extractor.extract_from_csv(str(cand)))
                elif ext == ".xml":
                    tables = list(extractor.extract_from_xml(str(cand)))

                if tables:
                    rows = list(generator.process_rows(extractor.map_and_clean(tables)))
                    if len(rows) > len(best_gen_rows):
                        best_gen_rows = rows
                        best_candidate = cand
            except Exception:
                continue

        gen_regs = {r["Info2"].split("_")[0]: r for r in best_gen_rows if "Info2" in r}
        gen_count = len(gen_regs)

        val_addrs = set(val_regs.keys())
        gen_addrs = set(gen_regs.keys())
        common = val_addrs.intersection(gen_addrs)

        type_match = 0
        coef_match = 0
        name_match = 0

        diff_sample = []

        for addr in common:
            vr = val_regs[addr]
            gr = gen_regs[addr]

            v_type = vr.get("type", "").strip().upper()
            g_type = gr.get("Info3", "").strip().upper()
            if v_type == g_type:
                type_match += 1

            v_c = vr.get("coef_a", "").strip()
            g_c = gr.get("CoefA", "").strip()
            try:
                if abs(float(v_c) - float(g_c)) < 1e-4:
                    coef_match += 1
            except (ValueError, TypeError):
                pass

            v_name = vr.get("name", "").strip().lower()
            g_name = gr.get("Name", "").strip().lower()
            if v_name == g_name or v_name in g_name or g_name in v_name:
                name_match += 1

            if len(diff_sample) < 3 and (
                v_type != g_type or abs(float(v_c or 0) - float(g_c or 0)) >= 1e-4
            ):
                diff_sample.append(
                    {
                        "addr": addr,
                        "rd_name": vr.get("name"),
                        "gen_name": gr.get("Name"),
                        "rd_type": v_type,
                        "gen_type": g_type,
                        "rd_coef": v_c,
                        "gen_coef": g_c,
                    }
                )

        match_count = len(common)
        if match_count > 0:
            category_stats[cat]["matched_files"] += 1
            category_stats[cat]["gen_total_regs"] += gen_count
            category_stats[cat]["common_regs"] += match_count
            category_stats[cat]["type_matches"] += type_match
            category_stats[cat]["coef_matches"] += coef_match
            category_stats[cat]["name_matches"] += name_match

        pct = (match_count / val_count * 100.0) if val_count else 0.0
        file_info = {
            "name": df.name,
            "status": "MATCH" if match_count > 0 else "NO_OVERLAP",
            "val_count": val_count,
            "gen_count": gen_count,
            "common": match_count,
            "match_pct": round(pct, 1),
            "source": best_candidate.name if best_candidate else None,
            "type_match_pct": round(type_match / match_count * 100.0, 1) if match_count else 0.0,
            "coef_match_pct": round(coef_match / match_count * 100.0, 1) if match_count else 0.0,
            "name_match_pct": round(name_match / match_count * 100.0, 1) if match_count else 0.0,
            "diff_sample": diff_sample,
        }
        category_stats[cat]["files_detail"].append(file_info)
        if match_count > 0:
            equipment_matches.append(file_info)

    out = {
        "categories": dict(category_stats),
        "equipment_matches": sorted(equipment_matches, key=lambda x: x["common"], reverse=True),
    }

    with open("scratch/rd_exact_metrics.json", "w", encoding="utf-8") as f:
        json.dump(out, f, indent=2, ensure_ascii=False)

    print("SUCCESS: Exact metrics saved to scratch/rd_exact_metrics.json")


if __name__ == "__main__":
    run_deep_comparison()
