import sys
import csv
import glob
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from audit_and_enrich_tags_strictly import resolve_sunpm_tag_strictly

csv_files = glob.glob('generated_definitions/*.csv')
unresolved_names = Counter()

for fpath in csv_files:
    with open(fpath, encoding='utf-8', errors='ignore') as f:
        reader = csv.reader(f, delimiter=';')
        header = next(reader, None)
        for row in reader:
            if not row or len(row) < 7:
                continue
            name = row[5].strip()
            raw_tag = row[6].strip()
            resolved = resolve_sunpm_tag_strictly(name, raw_tag)
            if not resolved:
                unresolved_names[name] += 1

print("Top 60 most frequent unresolved variable names:")
for name, cnt in unresolved_names.most_common(60):
    print(f"  {cnt:5d} : \"{name}\"")
