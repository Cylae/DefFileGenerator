r"""
Full Dataset Knowledge Mining and Verification Script.
Relaunches complete training across the entire dataset:
- \\WDN-FILES.WEBDYN.COM\Support\Equipementiers (recursive)
- Definitions/
- generated_definitions/

Constraints:
- Strictly READ-ONLY on all dataset files (zero modification to files in dataset).
- Preserves all elements already present in the learned dataset knowledge.
"""

import csv
import json
import os
import sys
import time
from pathlib import Path

# Add project root
WORKSPACE_DIR = Path(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
sys.path.insert(0, str(WORKSPACE_DIR))

from DefFileGenerator.knowledge_miner import (  # noqa: E402
    DEFAULT_DATASET_PATH,
    DatasetKnowledge,
    clean_variable_key,
    is_canonical_sunpm_tag,
    load_dataset_knowledge,
    normalize_tag_to_sunpm,
)

print("=" * 80)
print("RELAUNCHING COMPLETE DATASET KNOWLEDGE TRAINING")
print("Constraint: Strictly read-only on dataset files, preserving existing knowledge mappings")
print("=" * 80)

# Step 1: Load existing dataset knowledge
existing_knowledge = load_dataset_knowledge(DEFAULT_DATASET_PATH)
if existing_knowledge is None:
    print("No previous knowledge found, initializing new model.")
    knowledge = DatasetKnowledge()
    existing_keys = set()
else:
    knowledge = existing_knowledge
    existing_keys = set(knowledge.name_to_tag.keys())
    print(
        f"Loaded existing knowledge with {len(existing_keys):,} mappings, {knowledge.total_definitions_mined:,} definitions, {knowledge.total_registers_mined:,} registers."
    )

# Step 2: Resolve dataset directories
share_dir = Path(r"\\WDN-FILES.WEBDYN.COM\Support\Equipementiers")
dataset_dirs = [
    share_dir,
    WORKSPACE_DIR / "Definitions",
    WORKSPACE_DIR / "generated_definitions",
]

all_csv_files: list[Path] = []
seen_filepaths: set[str] = set()

t_scan = time.time()
for d in dataset_dirs:
    if not d.is_dir():
        print(f"Directory not found (skipping): {d}")
        continue
    files = [p for p in d.rglob("*.csv") if not p.name.startswith("~$")]
    for f in files:
        canonical_path = str(f.resolve()).lower()
        if canonical_path not in seen_filepaths:
            seen_filepaths.add(canonical_path)
            all_csv_files.append(f)

print(
    f"Discovered {len(all_csv_files):,} unique CSV files to train from across all dataset sources in {time.time() - t_scan:.2f}s."
)

# Step 3: Train strictly read-only
new_mappings_added = 0
mined_definitions = 0
mined_registers = 0

t_train = time.time()
for _idx, p in enumerate(all_csv_files, 1):
    try:
        # Strictly open in read-only mode with errors='ignore'
        with open(p, encoding="utf-8-sig", errors="ignore") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader, None)
            if not header or len(header) < 2 or not any(header):
                continue
            mined_definitions += 1

            for row in reader:
                if not row or len(row) < 7:
                    continue
                mined_registers += 1
                dtype = row[3].strip() if len(row) > 3 else ""
                name = row[5].strip() if len(row) > 5 else ""
                tag = row[6].strip() if len(row) > 6 else ""
                unit = row[9].strip() if len(row) > 9 else ""
                action = row[10].strip() if len(row) > 10 else ""

                if not name or not tag:
                    continue

                norm_tag = normalize_tag_to_sunpm(tag)
                if not is_canonical_sunpm_tag(norm_tag):
                    continue

                key = clean_variable_key(name)
                if not key:
                    continue

                # Preserve elements already present in dataset knowledge
                if key not in existing_keys:
                    knowledge.name_to_tag[key] = norm_tag
                    existing_keys.add(key)
                    new_mappings_added += 1

                # Update empirical distributions
                knowledge.tag_counts[norm_tag] += 1
                if unit:
                    knowledge.metric_units[norm_tag][unit.strip()] += 1
                if dtype:
                    knowledge.metric_types[norm_tag][dtype.strip().upper()] += 1
                if action:
                    knowledge.metric_actions[norm_tag][action.strip()] += 1

    except Exception:
        continue

train_duration = time.time() - t_train
knowledge.total_definitions_mined = max(knowledge.total_definitions_mined, mined_definitions)
knowledge.total_registers_mined = max(knowledge.total_registers_mined, mined_registers)

print(f"\nTraining completed in {train_duration:.2f}s:")
print(f"  Definitions scanned (read-only): {mined_definitions:,}")
print(f"  Registers scanned: {mined_registers:,}")
print(f"  New variable mappings added: {new_mappings_added:,}")
print(f"  Total variable-to-tag mappings: {len(knowledge.name_to_tag):,}")
print(f"  Unique canonical SunPM tags represented: {len(knowledge.tag_counts):,}")

print("\nTop 25 SunPM tags in trained dataset knowledge:")
for t, c in knowledge.tag_counts.most_common(25):
    print(f"  {t:<28}: {c:,} occurrences")

# Step 4: Persist the trained model into dataset_knowledge.json
output_json = DEFAULT_DATASET_PATH
with open(output_json, "w", encoding="utf-8") as f:
    json.dump(knowledge.to_dict(), f, indent=2)

print(
    f"\nSuccessfully saved trained knowledge model to: {output_json} ({output_json.stat().st_size:,} bytes)"
)
print("=" * 80)
