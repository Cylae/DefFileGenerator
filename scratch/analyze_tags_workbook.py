#!/usr/bin/env python3
"""Deep extraction and analysis of SunPM_DEfFile_Tags_v1.xlsx."""

import json
from pathlib import Path

import openpyxl


def analyze():
    wb = openpyxl.load_workbook("SunPM_DEfFile_Tags_v1.xlsx", data_only=True)

    results = {
        "sheets": wb.sheetnames,
        "categories": {},
        "tag_dictionary": {},
        "previous_tags_mapping": {},
        "synaptiq_indicators": [],
    }

    # 1. Sheet by sheet extraction
    # Standard tag sheets: TAG_INVERTER, TAG_METER, TAG_BATTERY, TAG_SENSOR, TAG_DISPLAY, TAG LOGGER, TAG_PLANT
    tag_sheets = [
        ("TAG_INVERTER", "Inverter"),
        ("TAG_METER", "Meter"),
        ("TAG_BATTERY", "Battery"),
        ("TAG_SENSOR", "Sensor"),
        ("TAG_DISPLAY", "Display"),
        ("TAG LOGGER", "Logger"),
        ("TAG_PLANT", "Plant"),
    ]

    for sheet_name, category in tag_sheets:
        ws = wb[sheet_name]
        cat_tags = []
        if sheet_name == "TAG_BATTERY":
            # Just single column of tags
            for r in range(1, ws.max_row + 1):
                val = ws.cell(r, 1).value
                if val:
                    cat_tags.append(
                        {
                            "tag": str(val).strip(),
                            "previous_tags": [],
                            "description": f"Battery {val}",
                            "mvp": True,
                            "category": category,
                        }
                    )
        else:
            header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
            # Locate columns
            tag_idx = 1
            prev_idx = None
            desc_idx = None
            mvp_idx = None

            for i, h in enumerate(header, start=1):
                h_str = str(h).lower() if h else ""
                if "tag" in h_str and "previous" not in h_str:
                    tag_idx = i
                elif "previous" in h_str:
                    prev_idx = i
                elif "desc" in h_str:
                    desc_idx = i
                elif "mvp" in h_str:
                    mvp_idx = i

            for r in range(2, ws.max_row + 1):
                tag_val = ws.cell(r, tag_idx).value
                if not tag_val:
                    continue
                tag_str = str(tag_val).strip()
                if tag_str.lower() in ("tags", "monitoring", "none"):
                    continue

                prev_val = ws.cell(r, prev_idx).value if prev_idx else None
                desc_val = ws.cell(r, desc_idx).value if desc_idx else None
                mvp_val = ws.cell(r, mvp_idx).value if mvp_idx else None

                prev_list = []
                if prev_val:
                    # Previous tags can be separated by '/', ';', '-'
                    raw_prev = str(prev_val)
                    for sep in ["/", ";", ","]:
                        raw_prev = raw_prev.replace(sep, "|")
                    prev_list = [p.strip() for p in raw_prev.split("|") if p.strip()]

                cat_tags.append(
                    {
                        "tag": tag_str,
                        "previous_tags": prev_list,
                        "description": str(desc_val).strip() if desc_val else "",
                        "mvp": str(mvp_val).lower() in ("x", "1", "true") if mvp_val else False,
                        "category": category,
                    }
                )

        results["categories"][category] = cat_tags

    # 2. Extract 3E SynaptiQ indicators
    ws_3e = wb["3E-input"]
    for r in range(2, ws_3e.max_row + 1):
        ind = ws_3e.cell(r, 1).value
        if not ind:
            continue
        disp_name = ws_3e.cell(r, 2).value
        desc = ws_3e.cell(r, 3).value
        dev_level = ws_3e.cell(r, 4).value
        unit = ws_3e.cell(r, 5).value
        priority = ws_3e.cell(r, 6).value
        inv_tag = ws_3e.cell(r, 7).value
        met_tag = ws_3e.cell(r, 8).value
        bat_tag = ws_3e.cell(r, 9).value
        sen_tag = ws_3e.cell(r, 10).value

        results["synaptiq_indicators"].append(
            {
                "indicator": str(ind).strip(),
                "display_name": str(disp_name).strip() if disp_name else "",
                "description": str(desc).strip() if desc else "",
                "device_level": str(dev_level).strip() if dev_level else "",
                "unit": str(unit).strip() if unit else "",
                "priority": priority,
                "tags": {
                    "Inverter": str(inv_tag).strip() if inv_tag else None,
                    "Meter": str(met_tag).strip() if met_tag else None,
                    "Battery": str(bat_tag).strip() if bat_tag else None,
                    "Sensor": str(sen_tag).strip() if sen_tag else None,
                },
            }
        )

    # Save to scratch
    out_file = Path("scratch/tags_workbook_analysis.json")
    out_file.write_text(json.dumps(results, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Analysis saved to {out_file}")

    # Print summary statistics
    for cat, tags in results["categories"].items():
        mvp_count = sum(1 for t in tags if t["mvp"])
        print(f"Category {cat:10}: {len(tags):2} tags (MVP: {mvp_count:2})")
    print(f"SynaptiQ indicators: {len(results['synaptiq_indicators'])}")


if __name__ == "__main__":
    analyze()
