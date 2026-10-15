import csv
from pathlib import Path
from DefFileGenerator.knowledge_miner import auto_fix_definition_file, normalize_tag_to_sunpm, infer_sunpm_tag

GEN_DIR = Path("generated_definitions")
files = list(GEN_DIR.glob("*.csv"))

total_normalized = 0
total_inferred = 0
total_dedup = 0

print(f"Applying SunPM Tag Enrichment & Normalization across {len(files)} definition files...")

for idx, f in enumerate(files, 1):
    try:
        counts = auto_fix_definition_file(
            input_path=f,
            output_path=f,
            auto_tag=True,
            normalize_legacy=True,
            deduplicate=True,
        )
        total_normalized += counts["tags_normalized"]
        total_inferred += counts["tags_inferred"]
        total_dedup += counts["duplicates_removed"]
    except Exception as e:
        print(f"Error on {f.name}: {e}")

print(f"Done!")
print(f"Legacy tags normalized to new SunPM canonical tags: {total_normalized:,}")
print(f"New SunPM tags inferred and added: {total_inferred:,}")
print(f"Duplicates removed: {total_dedup:,}")
