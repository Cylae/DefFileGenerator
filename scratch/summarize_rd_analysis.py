import json

with open('scratch/rd_deep_analysis_result.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

cats = data['categories']
print('=== CATEGORY SUMMARY ===')
total_rd_all = 0
total_gen_all = 0
total_common_all = 0
matched_files_all = 0
total_files_all = 0

for cat, s in sorted(cats.items()):
    total_files_all += s['total_files']
    matched_files_all += s['matched_files']
    total_rd_all += s['rd_total_regs']
    total_gen_all += s['gen_total_regs']
    total_common_all += s['common_regs']
    print(f"{cat}: files={s['matched_files']}/{s['total_files']}, rd_regs={s['rd_total_regs']}, gen_regs={s['gen_total_regs']}, common_regs={s['common_regs']}")

print(f"\nTOTAL ALL CATEGORIES: {matched_files_all}/{total_files_all} files, {total_common_all} common registers out of {total_rd_all} R&D registers.")

print('\n=== DETAILED EQUIPMENT SAMPLE (Top 30 by common registers) ===')
details = sorted(data['detailed_equipment'], key=lambda x: x['common'], reverse=True)
for d in details[:30]:
    tm = int(d['type_match_ratio'] * 100)
    cm = int(d['coef_match_ratio'] * 100)
    pct = d['match_pct']
    print(f"- {d['name']}: R&D={d['val_count']} | Gen={d['gen_count']} | Common={d['common']} ({pct}%) | TypeMatch={tm}% | CoefMatch={cm}% | Src={d['source']}")

print('\n=== TYPICAL DIFFERENCES DETECTED ===')
diff_count = 0
for d in details:
    if d.get('diff_sample'):
        print(f"\nDiscrepancy in {d['name']}:")
        for diff in d['diff_sample']:
            print(f"  Addr {diff['addr']}: RD Name='{diff['rd_name']}', Gen Name='{diff['gen_name']}' | Type RD={diff['rd_type']} vs Gen={diff['gen_type']} | CoefA RD={diff['rd_coef']} vs Gen={diff['gen_coef']}")
            diff_count += 1
            if diff_count >= 15:
                break
    if diff_count >= 15:
        break
