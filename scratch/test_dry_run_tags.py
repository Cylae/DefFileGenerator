import sys
from pathlib import Path
import csv
import glob
from collections import Counter

sys.path.insert(0, str(Path(__file__).parent))
from audit_and_enrich_tags_strictly import resolve_sunpm_tag_strictly, is_valid_official_tag

csv_files = glob.glob('generated_definitions/*.csv')
print(f"Analyzing {len(csv_files)} files in generated_definitions...")

total_registers = 0
initial_empty = 0
initial_canonical = 0
initial_non_canonical = 0

final_empty = 0
final_canonical = 0
final_non_canonical = 0

purged_samples = Counter()
enriched_samples = Counter()
tag_distribution = Counter()

for fpath in csv_files:
    with open(fpath, encoding='utf-8', errors='ignore') as f:
        reader = csv.reader(f, delimiter=';')
        header = next(reader, None)
        for row in reader:
            if not row or len(row) < 7:
                continue
            total_registers += 1
            name = row[5].strip()
            raw_tag = row[6].strip()
            
            if not raw_tag:
                initial_empty += 1
            elif is_valid_official_tag(raw_tag):
                initial_canonical += 1
            else:
                initial_non_canonical += 1
                
            resolved = resolve_sunpm_tag_strictly(name, raw_tag)
            if not resolved:
                final_empty += 1
                if raw_tag and not is_valid_official_tag(raw_tag):
                    purged_samples[raw_tag] += 1
            elif is_valid_official_tag(resolved):
                final_canonical += 1
                tag_distribution[resolved] += 1
                if not raw_tag:
                    enriched_samples[(name, resolved)] += 1
            else:
                final_non_canonical += 1

print(f"Total registers: {total_registers}")
print(f"Initial: {initial_canonical} canonical, {initial_non_canonical} non-canonical, {initial_empty} empty")
print(f"Final:   {final_canonical} canonical, {final_non_canonical} non-canonical, {final_empty} empty")
print("\nTop 20 purged pseudo-tags:")
for t, cnt in purged_samples.most_common(20):
    print(f"  {t}: {cnt}")
print("\nTop 20 newly enriched variables:")
for (name, t), cnt in enriched_samples.most_common(20):
    print(f"  \"{name}\" -> {t}: {cnt}")
