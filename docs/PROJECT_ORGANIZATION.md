# Project Organization & Technical Architecture

This document defines the canonical organization, file format specifications, generation workflows, and validation contracts for **DefFileGenerator**.

---

## 1. Canonical Per-Manufacturer & Per-Range Layout

Every definition catalog and staging repository follows a strict three-tier hierarchy organized by **Manufacturer** and **Equipment Range (Series)**:

```text
<root>/                                  # production/, Definitions/, generated_definitions/, tagged_definitions/
└── <Manufacturer>/                      # Canonical PascalCase manufacturer (e.g., FoxESS, Huawei, SMA, ABB)
    └── <Range>/                         # Equipment family / series (e.g., P3_H3_Pro, SUN2000, Sunny_Tripower)
        ├── <PREFIX>_<Manufacturer>_<Model>.csv   # WebdynSunPM definition file (.csv)
        ├── <PREFIX>_<Manufacturer>_<Model>.ini   # WebdynSun Classic definition file (.ini)
        └── sources/                              # Manufacturer source documentation & traceability
            ├── <manual>.pdf / <regmap>.xlsx / <spec>.xml
            └── README.md                         # Range provenance, file index, and SHA-256 digests
```

### Key Architectural Rules

1. **Dual Target Parity in Every Range Directory**:
   - **WebdynSunPM (`.csv`)**: Semicolon-delimited 11-column CSV, pure UTF-8 without BOM, CRLF (`\r\n`) line endings, sequential 1-based row indices, normalized `Modbus;` protocol header, and 6-decimal numeric scaling coefficients (`CoefA`, `CoefB`).
   - **WebdynSun Classic (`.ini`)**: Native WebdynSun Gateway v2.0 INI configuration containing `Modbus_RequestsTables={...}` (10 fields per request, contiguous reads bounded to `<= 32` registers, writable parameters mapped with FC 5/6/16) and `Modbus_VariablesTables={...}` (12 fields per variable, exact `varType` numeric codes `1..9`, explicit signedness `1=signed` / `2=unsigned`, 1-based `varPosition` within the Modbus response packet, and mapped `varAction`).
2. **Co-located `sources/` Subdirectory**:
   - Every `<Range>/` directory contains a `sources/` subdirectory holding the exact manufacturer manuals, register workbooks (PDF, XLSX, XML, JSON, CSV), and an English `README.md` with SHA-256 checksums linking each source document to the generated `.csv` and `.ini` definitions.
   - Any recursive scanner (`iter_definition_csvs`, `find_definition_file`, `verify_file_pillars`, `DuplicateDetector`, `DefinitionVersionControl`, FastAPI Web Studio, and `validate_production.py`) explicitly excludes `sources/` paths when discovering Webdyn definition CSVs so raw manufacturer source CSVs are never misclassified as Webdyn gateway definitions.

---

## 2. FoxESS Ecosystem Layout (`production/FoxESS/`)

Special care applies to the **FoxESS** catalog, which separates each hardware series and protocol generation into its own range directory with verified WebdynSunPM (`.csv`), WebdynSun (`.ini`), and official protocol sources:

| Range Directory | Equipment Scope | WebdynSunPM (`.csv`) | WebdynSun (`.ini`) | Source Documents (`sources/`) |
|---|---|---|---|---|
| `P3_H3_Pro/` | Combined P3 Pro & H3 Pro (Inverter, Meter, Battery, All-In-One) | `*_FoxESS_P3_H3_PRO*.csv` | `*.ini` + `REF_FoxESS_P3_H3_PRO_INVERTER.ini` | `FoxESS Modbus Protocol--20260701 (V1.05.05.00).pdf`, `FoxESS-P3-5.0-SH-P3-15.0-SH-Oswiadczenie.pdf` |
| `P3_Pro/` | P3 Pro dedicated definitions | `*_FoxESS_P3_Pro.csv` | `*.ini` | `FoxESS Modbus Protocol--20260701 (V1.05.05.00).pdf` |
| `H3_Pro/` | H3 Pro dedicated definitions | `*_FoxESS_H3_Pro.csv` | `*.ini` | `FoxESS Modbus Protocol--20260701 (V1.05.00.00).xlsx/.pdf` |
| `G_MAX/` | G-MAX Storage, BMS, EMS & ModbusTCP | `*_FoxESS_G-MAX*.csv` | `*.ini` | `G-MAX Communication Protocol modbusTcp-EN-20250425(1).xlsx` |
| `R_Series/` | R-Series Inverters & Batteries | `*_FoxESS_RSeries*.csv` | `*.ini` + `REF_INV_FOXESS_RSERIESINVERTER.ini` | `FoxESS RSeries xml file.xml`, `FR-R-Series-Datasheet-4.20.pdf`, `RSeries - Sunspec_Modbus_FOXESS_V1.04..1.06.pdf` |
| `R_V_H3plus/` | R/V/H3plus/GmaxPlus SunSpec maps | `*_FoxESS_*H3plus*.csv` | `*.ini` | `RVH3plus- Sunspec_Modbus_FOXESS_V1.07.pdf`, `R&V&H3plus&GmaxPlus-_Sunspec_Modbus_FOXESS_V1.08.pdf` |

---

## 3. Repository Directories Overview

```text
DefFileGenerator/
├── production/                   # Master validated library (<Manufacturer>/<Range>/{.csv, .ini, sources/})
├── Definitions/                  # Legacy compatibility mirror using identical <Manufacturer>/<Range>/ layout
├── generated_definitions/        # Parser staging output using <Manufacturer>/<Range>/ layout
├── tagged_definitions/           # SunPM semantically enriched mirror using <Manufacturer>/<Range>/ layout
├── sources/                      # Cryptographically sealed master manual archive (`MANIFEST.json`)
├── DefFileGenerator/             # Core Python package
│   ├── def_gen.py                # Core WebdynSunPM generator, type/address normalization, and validation
│   ├── webdynsun_ini.py          # Official WebdynSun Classic (.ini) converter & structural validator
│   ├── extractor.py              # Multi-format PDF/XLSX/CSV/XML/JSON Modbus register extractor
│   ├── production_governor.py    # 5-Pillar compliance validator (naming, UTF-8 no BOM, CRLF, FCs, tags)
│   ├── io_utils.py               # Atomic staged writes, recursive find_definition_file() & iter_definition_csvs()
│   ├── knowledge_miner.py        # Canonical WebdynSunPM semantic tag inference & auto-tagging
│   ├── duplicate_detector.py     # SHA-256, structural, and alias duplicate analyzer
│   ├── version_manager.py        # Content-addressable definition snapshot versioning
│   ├── *_definitions_engine.py   # 14 manufacturer-specific generation engines
│   ├── sunspec/                  # SunSpec PICS Excel parser and DER 700-series model generator
│   ├── discovery/                # Autonomous Modbus discovery scaffold (Phase B specification target)
│   └── tests/                    # Automated unit, contract, security, and production certification suites
├── web/                          # FastAPI Web Studio (`web/app.py`, `web/static/`)
├── scripts/                      # Maintenance, validation (`validate_production.py`), and reorganization CLIs
└── docs/                         # Technical documentation, architecture specs, and audit ledgers
```

---

## 4. Target Format Specifications

### 4.1 WebdynSunPM (`.csv`) — 5 Apollo Pillars

1. **Filename Convention**: `<PREFIX>_<Manufacturer>_<Model>.csv` where `<PREFIX>` is in `{INV, MET, SEN, BAT, COM, OTH}`.
2. **Binary Hygiene**: Strict UTF-8 without BOM (`\xef\xbb\xbf` forbidden) and CRLF (`\r\n`) line terminators.
3. **Header**: Line 1 has 11 semicolon-delimited columns starting with `Modbus;<Category>;<Manufacturer>;<Model>;;;;;;;`.
4. **Row Invariants**: Sequential 1-based `Index` matching row order, Modbus read function codes `{1, 2, 3, 4}`, bounded decimal addresses `0..65535` (or `addr_byteLen` for `STRING`/`RAW`, `addr_bit_len` for `BITS`), 6-decimal `CoefA`/`CoefB`, valid Webdyn action codes `{0..12}`, and unique single-instance physical SunPM tags.

### 4.2 WebdynSun Classic (`.ini`) — Gateway Specification v2.0

Generated and verified by [`DefFileGenerator/webdynsun_ini.py`](../DefFileGenerator/webdynsun_ini.py):

- **`Modbus_RequestsTables`**:
  - Fields: `reqIndex;reqName;reqReadFctCode;reqWriteFctCode;reqStartRegister;reqNbRegisters;reqEnableReading;reqEnableWriting;reqOption1;reqOption2`
  - Contiguous read registers are grouped up to `MAX_WEBDYNSUN_READ_REGS = 32` per request.
  - Writable parameters (`action == 1`) receive dedicated write-enabled requests (`writeFctCode` = `5` for coils, `6` for single holding register, `16` for multi-register values).
- **`Modbus_VariablesTables`**:
  - Fields: `varIndex;varReqIndex;varName;varType;varSigned;varPosition;varOption1;varOption2;coeffA;coeffB;unit;varAction`
  - `varType` codes: `1` (Bit), `2` (Byte), `3` (Word), `4` (Reversed Word), `5` (Double Word), `6` (Reversed Double Word), `7` (Float32), `8` (Reversed Float32), `9` (String).
  - `varSigned`: `1` (signed) or `2` (unsigned).
  - `varPosition`: 1-based register (or bit) offset inside the parent request packet (`1 + (reg_addr - req_start)`).
  - Unsupported 64-bit (`U64`/`I64`/`F64`) or network (`IP`/`MAC`) registers are explicitly documented in `# NOTE:` header comments rather than silently truncated.

---

## 5. Validation & Maintenance Commands

```bash
# Reorganize any new definition files into <Manufacturer>/<Range>/{.csv, .ini, sources/}
python scripts/reorganize_repo.py

# Validate all WebdynSunPM (.csv), WebdynSun (.ini), and sources/ directories in production/
python scripts/validate_production.py

# Run full automated verification suite
pytest -q
```
