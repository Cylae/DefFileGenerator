# Repository Directory Taxonomy

This document summarizes the canonical folder layout and governance rules of **DefFileGenerator**.
For the full technical specification (WebdynSunPM `.csv`, WebdynSun Classic `.ini`, FoxESS ranges, and `sources/` traceability), see:
-> **[`docs/PROJECT_ORGANIZATION.md`](PROJECT_ORGANIZATION.md)**

---

## Functional Domains Overview

```text
DefFileGenerator/
├── production/                      # Master certified catalog (<Manufacturer>/<Range>/{.csv, .ini, sources/})
│   ├── CATALOG.json                 # Machine-readable catalog index
│   ├── SOURCES.md                   # Source traceability matrix
│   └── README.md                    # Production governance guide
├── Definitions/                     # Historical compatibility mirror (<Manufacturer>/<Range>/{.csv, .ini, sources/})
├── generated_definitions/           # Parser raw staging (<Manufacturer>/<Range>/{.csv, .ini, sources/})
├── tagged_definitions/              # Semantically enriched staging (<Manufacturer>/<Range>/{.csv, .ini, sources/})
├── sources/                         # Sealed raw source manuals (MANIFEST.json with SHA-256)
├── DefFileGenerator/                # Core Python package (def_gen, webdynsun_ini, extractor, governor, sunspec, discovery, tests)
├── web/                             # FastAPI Web Studio backend & frontend
├── scripts/                         # Validation, reorganization, and maintenance CLI tools
└── docs/                            # Technical documentation (`PROJECT_ORGANIZATION.md`, `ARCHITECTURE.md`, `audit/`)
```

## Governance Rules

1. **Per-Manufacturer & Per-Range Structure**: Every equipment range resides in `<root>/<Manufacturer>/<Range>/` with both WebdynSunPM (`.csv`) and WebdynSun Classic (`.ini`) definition files plus a `sources/` subdirectory.
2. **Read-Only Network Share**: `\\WDN-FILES.WEBDYN.COM\Support\Equipementiers` is treated strictly as a read-only reference dataset unless explicit user write authorization is granted.
3. **5-Pillar + `.ini` Certification**: Every definition file under `production/` must pass `python scripts/validate_production.py` (CSV 5 Pillars + WebdynSun `.ini` structure + range `sources/`).
