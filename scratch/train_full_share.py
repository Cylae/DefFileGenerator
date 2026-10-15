#!/usr/bin/env python3
"""
Full Dataset Knowledge Mining and Verification Script.
Mines \\WDN-FILES.WEBDYN.COM\Support\Equipementiers
"""

import csv
import json
import os
import sys
import time
from pathlib import Path

# Add project root
sys.path.insert(0, os.path.abspath("."))

from DefFileGenerator.knowledge_miner import (
    DatasetKnowledge,
    is_canonical_sunpm_tag,
    normalize_tag_to_sunpm,
)

share_dir = Path(r"\\WDN-FILES.WEBDYN.COM\Support\Equipementiers")
knowledge = DatasetKnowledge()

print(f"Mining dataset knowledge from: {share_dir}...")
start = time.time()
csv_files = list(share_dir.rglob("*.csv"))
print(f"Discovered {len(csv_files)} CSV files in {time.time() - start:.2f}s.")

parse_start = time.time()
for p in csv_files:
    try:
        with open(p, encoding="utf-8-sig", errors="ignore") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader, None)
            if not header or len(header) < 4:
                continue
            knowledge.total_definitions_mined += 1
            for row in reader:
                if not row or len(row) < 7:
                    continue
                knowledge.total_registers_mined += 1
                dtype = row[3].strip() if len(row) > 3 else ""
                name = row[5].strip() if len(row) > 5 else ""
                tag = row[6].strip() if len(row) > 6 else ""
                unit = row[9].strip() if len(row) > 9 else ""
                action = row[10].strip() if len(row) > 10 else ""

                norm_tag = normalize_tag_to_sunpm(tag)
                if not is_canonical_sunpm_tag(norm_tag):
                    continue

                knowledge.learn_entry(
                    name=name,
                    tag=norm_tag,
                    dtype=dtype,
                    unit=unit,
                    action=action,
                )
    except Exception:
        continue

elapsed = time.time() - parse_start
print(f"Mined {knowledge.total_definitions_mined} definitions, {knowledge.total_registers_mined:,} registers in {elapsed:.2f}s.")
print(f"Learned variable-to-tag mappings: {len(knowledge.name_to_tag):,}")
print(f"Top 20 SunPM tags learned:")
for t, c in knowledge.tag_counts.most_common(20):
    print(f"  {t:<25}: {c} occurrences")

# Save full mined knowledge to DefFileGenerator/dataset_knowledge.json
output_json = Path("DefFileGenerator/dataset_knowledge.json")
with open(output_json, "w", encoding="utf-8") as f:
    json.dump(knowledge.to_dict(), f, indent=2)
print(f"\nPersisted learned model to {output_json} ({output_json.stat().st_size:,} bytes)")
