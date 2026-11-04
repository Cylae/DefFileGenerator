import sys
import csv
import glob
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from audit_and_enrich_tags_strictly import resolve_sunpm_tag_strictly

csv_files = glob.glob('generated_definitions/*.csv')
telemetry_unresolved = Counter()

keywords = ['power', 'energy', 'voltage', 'current', 'temp', 'frequency', 'puissance', 'energie', 'tension', 'courant', 'watt', 'volt', 'amp', 'hz']

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
                name_l = name.lower()
                if any(k in name_l for k in keywords):
                    telemetry_unresolved[name] += 1

print(f"Total distinct unresolved with telemetry keywords: {len(telemetry_unresolved)}")
print("\nTop 80 unresolved with telemetry keywords:")
for name, cnt in telemetry_unresolved.most_common(80):
    print(f"  {cnt:4d} : \"{name}\"")
