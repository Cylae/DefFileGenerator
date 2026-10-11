import json

with open("scratch/rd_exact_metrics.json", "r", encoding="utf-8") as f:
    data = json.load(f)

cats = data["categories"]
total_rd = sum(s["rd_total_regs"] for s in cats.values())
total_common = sum(s["common_regs"] for s in cats.values())
total_type = sum(s["type_matches"] for s in cats.values())
total_coef = sum(s["coef_matches"] for s in cats.values())
total_name = sum(s["name_matches"] for s in cats.values())

print("=== TOTAL AGGREGATES ===")
print(f"Total R&D registers: {total_rd}")
print(f"Total Common Modbus addresses extracted: {total_common}")
print(f"Type matches: {total_type} / {total_common} ({total_type / total_common * 100:.1f}%)")
print(f"CoefA matches: {total_coef} / {total_common} ({total_coef / total_common * 100:.1f}%)")
print(
    f"Name/Semantic matches: {total_name} / {total_common} ({total_name / total_common * 100:.1f}%)"
)

print("\n=== PER CATEGORY ===")
for cat, s in sorted(cats.items()):
    c = s["common_regs"]
    tm = (s["type_matches"] / c * 100) if c else 0
    cm = (s["coef_matches"] / c * 100) if c else 0
    nm = (s["name_matches"] / c * 100) if c else 0
    m_files = s["matched_files"]
    t_files = s["total_files"]
    rd_r = s["rd_total_regs"]
    print(
        f"{cat:15} | Files: {m_files}/{t_files} | Common: {c:4} / {rd_r:4} | Types: {tm:5.1f}% | CoefA: {cm:5.1f}% | Names: {nm:5.1f}%"
    )

print("\n=== TOP 20 EQUIPMENT COMPARISONS ===")
for eq in data["equipment_matches"][:20]:
    name = eq["name"]
    val_c = eq["val_count"]
    gen_c = eq["gen_count"]
    comm = eq["common"]
    pct = eq["match_pct"]
    tm = eq["type_match_pct"]
    cm = eq["coef_match_pct"]
    nm = eq["name_match_pct"]
    src = eq["source"]
    print(f"- {name}:")
    print(f"    R&D: {val_c} regs | Gen: {gen_c} regs | Common: {comm} ({pct}%)")
    print(f"    Concordance -> Types: {tm}% | CoefA: {cm}% | Noms: {nm}%")
    print(f"    Source: {src}")
