# DefFileGenerator — Developer Architecture & Onboarding Guide

<p align="center">
  <a href="README.md"><b>README</b></a> •
  <a href="docs/development.md"><b>Development</b></a> •
  <a href="docs/architecture.md"><b>Architecture</b></a> •
  <a href="docs/security.md"><b>Security</b></a> •
  <a href="docs/audit/VALIDATION.md"><b>Audit Ledger</b></a>
</p>

## Table of contents

1. [Executive Summary & Purpose](#1-executive-summary--purpose)
2. [System Architecture & Data Flow](#2-system-architecture--data-flow)
3. [WebdynSunPM Specification & Domain Invariants](#3-webdynsunpm-specification--domain-invariants)
4. [Key Implementation Details & Algorithms](#4-key-implementation-details--algorithms)
5. [Developer Recipes: How to Extend the Codebase](#5-developer-recipes-how-to-extend-the-codebase)
6. [Development & Quality Assurance Workflows](#6-development--quality-assurance-workflows)
7. [Apollo-Standard Autonomous Constitution & Ledgers](#7-apollo-standard-autonomous-constitution--ledgers)
8. [Standalone Windows Executables (.exe)](#8-standalone-windows-executables-exe)

Welcome to the **DefFileGenerator** project! This guide is written for software engineers maintaining, debugging, or extending this codebase. It documents the domain concepts, system architecture, data structures, invariants, and standard development workflows.

---

## 1. Executive Summary & Purpose

`DefFileGenerator` automates the conversion of vendor-specific Modbus documentation into syntactically valid, validated **WebdynSunPM** definition files (`.csv`).

Manufacturer register maps vary wildly in format:
- Different file types (Adobe PDF, Microsoft Excel, CSV, XML).
- Inconsistent column headers (`"Register"`, `"Addr (dec)"`, `"Point"`, `"Grandeur"`, `"Data Format"`).
- Diverse data types (`"uint16"`, `"FLOAT32 Big Endian"`, `"Bitfield16"`, `"STR20"`).
- Scaling variations (`"Gain: 10"` divisor vs. `"Multiplier: 0.1"` vs. `"Scale Factor: -1"`).

`DefFileGenerator` provides a format-agnostic extraction, semantic tagging, and normalization pipeline accessible via:
1. **Command Line Interface (`deffilegen`)** — Primary automation tool for pipelines and scripting.
2. **Industrial Web Studio & REST API (`web/app.py`)** — FastAPI server powering a 4-tab browser-based conversion, tagging, auditing, and equipment library studio.
3. **Windows 11 Desktop GUI (`DefFileGenerator/gui.py`)** — Modern CustomTkinter application with live register preview.
4. **Semantic Domain Engine (`DefFileGenerator/knowledge_miner.py`)** — Trained on 211,000+ registers for SunPM canonical tagging.
5. **Technical Audit & Report Generator (`DefFileGenerator/report_generator.py`)** — Markdown and JSON audit reporting.

---

## 2. System Architecture & Data Flow

The codebase strictly enforces a layered architecture where the core domain logic has zero dependency on UI or web frameworks:

```text
[ PDF / Excel / CSV / XML ]
            │
            ▼
┌────────────────────────────────────────────────────────┐
│  DefFileGenerator/extractor.py                         │
│  - Stream buffer ingestion (avoids OS file locks)      │
│  - Smart header detection & Modbus keyword scoring     │
│  - Multi-row header banner merging                     │
│  - Statistical column type inference fallback          │
│  - Exact, synonym, partial, and PDF fallback mapping   │
└───────────────────────────┬────────────────────────────┘
                            │ Yields intermediate register dicts
                            ▼
┌────────────────────────────────────────────────────────┐
│  DefFileGenerator/knowledge_miner.py                   │
│  - 211,000+ register training knowledge base           │
│  - Official SunPM legacy-to-canonical tag mapping     │
│  - Phase detection (_1, _2, _3) and unit disambiguation │
│  - Action code inference (4=Telemetry, 9=Alarm, 1=Cmd) │
└───────────────────────────┬────────────────────────────┘
                            │ Enriched register stream
                            ▼
┌────────────────────────────────────────────────────────┐
│  DefFileGenerator/def_gen.py                           │
│  - Data type normalization & endianness tagging        │
│  - Address boundary validation (0..65535)              │
│  - Bisect interval lookup and address overlap checks  │
│  - Linear coefficient computation (CoefA, CoefB)       │
│  - Unique variable tag synthesis & deduplication       │
│  - CSV injection neutralization                        │
│  - Atomic disk persistence with fsync                  │
└───────────────────────────┬────────────────────────────┘
                            │ Generates validated definition CSV
                            ▼
┌────────────────────────────────────────────────────────┐
│  Output Layer (Atomic Staging & fsync)                 │
│  - Writes to hidden staging file: .{name}.{uuid}.tmp   │
│  - Explicit flush and physical disk sync via os.fsync  │
│  - Atomic replacement via os.replace                   │
└───────────────────────────┬────────────────────────────┘
                            │
        ┌───────────────────┼───────────────────┬───────────────────┐
        ▼                   ▼                   ▼                   ▼
┌───────────────┐   ┌───────────────┐   ┌───────────────┐   ┌───────────────┐
│      CLI      │   │  FastAPI Web  │   │  Windows GUI  │   │ Report Engine │
│    main.py    │   │  web/app.py   │   │    gui.py     │   │ report_gen.py │
└───────────────┘   └───────────────┘   └───────────────┘   └───────────────┘
```

---

## 3. WebdynSunPM Specification & Domain Invariants

WebdynSunPM data loggers communicate with Modbus RTU/TCP inverters, meters, and weather stations. The generator writes semicolon-delimited (`;`) UTF-8-with-BOM definition CSV files; the validator also accepts supported legacy encodings.

### Line 1: Header Row
| Column Index | Field Name | Description | Example |
|---|---|---|---|
| 0 | `Protocol` | Communication protocol | `modbusRTU` or `modbusTCP` |
| 1 | `Category` | Equipment classification | `Inverter`, `Meter`, `Sensor` |
| 2 | `Manufacturer` | Equipment brand | `Huawei`, `ABB`, `SMA` |
| 3 | `Model` | Device model identifier | `SUN2000`, `STP5000` |
| 4 | `Forced writing code` | Optional forced write command | Empty or command string |
| 5–10 | Empty | Padding so the header has the same width as register rows | `;;;;;;` |

### Lines 2+: Register Rows
| Column Index | Field Name | Type | Description |
|---|---|---|---|
| 0 | `Index` | Integer | 1-based sequential line counter (`1`, `2`, `3`, ...) |
| 1 | `Info1` | Integer | Modbus function code category: `1`=Coil, `2`=Discrete Input, `3`=Holding Register, `4`=Input Register |
| 2 | `Info2` | String | Modbus address (e.g. `40001`, `40001_20` for strings, `40001_0_16` for bits) |
| 3 | `Info3` | String | Normalized Webdyn type: `U16`, `I16`, `U32_WB`, `F32_WB`, `STRING`, `RAW`, `BITS`, etc. |
| 4 | `Info4` | String | Reserved for future firmware expansion (always empty string `""`) |
| 5 | `Name` | String | Human-readable variable description (e.g. `"Active Power"`) |
| 6 | `Tag` | String | Unique variable identifier or canonical SunPM tag (`P_AC`, `active_power`) |
| 7 | `CoefA` | Float | Multiplier factor formatted to 6 decimal places ($a \cdot 10^{\text{scale}}$) |
| 8 | `CoefB` | Float | Offset addition bias formatted to 6 decimal places |
| 9 | `Unit` | String | Engineering measurement unit (`V`, `A`, `W`, `kW`, `Hz`, `°C`) |
| 10 | `Action` | String | Webdyn action code: `0` (reserve), `1` (command), `4` (telemetry), `9` (alarm), `10` (constant) |

### Scaling Polynomial Formula
Webdyn loggers compute the physical reading using the linear equation:
$$\text{Physical Value} = \text{CoefA} \times \text{Raw Register Value} + \text{CoefB}$$

In `DefFileGenerator`:
- `CoefA` = $\text{Factor} \times 10^{\text{ScaleFactor}}$ (default: `1.000000`)
- `CoefB` = $\text{Offset}$ (default: `0.000000`)

---

## 4. Key Implementation Details & Algorithms

### 4.1. Address Overlap Detection (`_check_address_overlap`)
Modbus addresses cannot overlap within the same `Info1` memory space (e.g. two 32-bit registers cannot share address `40001` and `40002`).

- **Interval Model**: Each register is treated as an interval $[\text{start}, \text{start} + \text{count} - 1]$.
- **Binary Search**: Uses `bisect.bisect_left` on sorted interval lists for $O(\log N)$ lookup.
- **Bit-Slice Exception**: Multiple `BITS` entries inside the same 16-bit word (e.g. `30001_0_1` and `30001_1_1`) do **not** trigger a collision if their bit ranges $[start, start + length - 1]$ are disjoint.

### 4.2. CSV Formula Injection Defense (`sanitize_csv_field`)
Spreadsheet applications execute formulas if a cell begins with `=, +, -, @, |, %`, or certain tab/whitespace triggers.
- Any cell starting with these triggers has an apostrophe (`'`) prepended.
- Finite numeric literals (e.g. `-10`, `+5.4`) are recognized and preserved without apostrophes.
- Multiline newlines (`\r`, `\n`) are replaced with single spaces to prevent CSV row desynchronization.

### 4.3. Safe Atomic File Operations
When writing output CSV files:
1. Data is written to a temporary sibling file (`.{name}.{uuid}.tmp`) in the destination directory.
2. Changes are flushed and synced to disk via `os.fsync(fd)`.
3. Configured generation validates the complete staging file before replacement; `os.replace` publishes the completed file atomically.

---

## 5. Developer Recipes: How to Extend the Codebase

### Recipe 1: Adding a New Synonym for a Column Header
Open `DefFileGenerator/extractor.py` and add your synonym (lowercase) to `Extractor.COLUMN_MAPPING`:
```python
"Factor": [
    "scale",
    "factor",
    "multiplier",
    "ratio",
    "mon_nouveau_synonyme",  # <--- Add here
],
```

### Recipe 2: Adding a New Data Type Synonym
Open `DefFileGenerator/def_gen.py` and append your compiled regex pattern to `TYPE_SYNONYMS_COMPILED`:
```python
TYPE_SYNONYMS_COMPILED: tuple[tuple[re.Pattern[str], str], ...] = (
    # ... existing patterns ...
    (re.compile(r"\bmon_type_custom\b", re.IGNORECASE), "U32"),
)
```

### Recipe 3: Adding a New File Format Extractor
1. In `DefFileGenerator/extractor.py`, implement an extraction method yielding generator streams:
```python
def extract_from_json(self, filepath: str) -> Iterator[Iterator[dict[str, Any]]]:
    def json_generator() -> Iterator[dict[str, Any]]:
        with open(filepath, "r", encoding="utf-8") as f:
            data = json.load(f)
            for item in data.get("registers", []):
                yield item

    yield json_generator()
```
2. Register the file extension in `ALLOWED_EXTENSIONS` in `DefFileGenerator/main.py`.
3. Dispatch the new parser in `_perform_extraction` in `main.py` and `generate_webdyn_def.py`.

### Recipe 4: Adding a New SunPM Canonical Tag Rule
Open `DefFileGenerator/knowledge_miner.py` and add the legacy or synonym mapping to `SUNPM_LEGACY_TO_CANONICAL`:
```python
SUNPM_LEGACY_TO_CANONICAL: dict[str, str] = {
    # ...
    "MonAncienSignal": "P_AC",
    "GridV_Phase1": "U_AC_1",
}
```

### Recipe 5: Generating Custom Technical Audit Reports
To programmatically generate a Markdown quality report from a directory of definition files:
```python
from DefFileGenerator.report_generator import generate_batch_report

markdown_content, stats = generate_batch_report(
    target_dir="Definitions",
    output_markdown="docs/AUDIT_REPORT.md",
    json_export_path="docs/metrics.json",
)
print(f"Audited {stats.total_files} definitions ({stats.valid_files} valid).")
```

### Recipe 6: Authoring an Industrial Equipment Generator
For complex multi-slave topologies (e.g. BESS, Central Inverters), follow the pattern in `DefFileGenerator/mars_generator.py`:
1. Define slave roles (e.g., Slave 1 Controller, Slave 247 EMS Overview, Slave N Racks).
2. Use `write_definition_file` with explicit validation against `Generator.validate_csv_detailed`.
3. Enforce UTF-8 BOM encoding and 11-column header alignment.

---

## 6. Development & Quality Assurance Workflows

All code must pass static typing, linting, security scanning, and automated tests before committing.

### Running Tests
```powershell
# Run the complete test suite with coverage report (1,090 tests, 90% coverage):
uv run pytest --cov=DefFileGenerator --cov=web

# Run fast targeted unit tests:
uv run pytest DefFileGenerator/tests/test_def_gen.py DefFileGenerator/tests/test_extractor.py
```

### Linting & Code Formatting
```powershell
# Check for lint violations:
uv run ruff check .

# Check formatting compliance:
uv run ruff format --check .

# Auto-format all code:
uv run ruff format .
```

### Static Type Checking (Strict Mypy)
```powershell
uv run mypy DefFileGenerator web
```

### Static Security Analysis (Bandit)
```powershell
uv run bandit -r DefFileGenerator web -ll
```

### Packaging & Build Validation
```powershell
uv build --offline
```

---

## 7. Apollo-Standard Autonomous Constitution & Ledgers

This repository is governed by the **Apollo-Standard Engineering Constitution** (`AGENTS.md`). All changes must be backed by empirical evidence, zero placeholders, and strict backward compatibility.

> [!IMPORTANT]
> **Mandatory User Rule for Shared Network Storage:**  
> Toujours demander confirmation explicite à l'utilisateur avant d'écrire, modifier, écraser ou supprimer tout fichier situé sur le partage réseau commun (`\\WDN-FILES.WEBDYN.COM`).

### Audit Ledgers in `docs/audit/`:
- **`PLAN.md`**: Milestone tracking and task completion status.
- **`REVIEW_MATRIX.md`**: LOC, risk classification, and test coverage per module.
- **`FINDINGS.md`**: Root-cause defect ledger (FIND-001 through FIND-008).
- **`VALIDATION.md`**: Output traces for all quality gates.
- **`PERFORMANCE.md`**: Collection benchmarks (2.50s test collection) and profiling.
- **`HANDOFF.md`**: Operational handoff instructions.

---

## 8. Standalone Windows Executables (.exe)

The repository provides an automated PyInstaller pipeline to compile self-contained 64-bit Windows executables requiring zero Python runtime on the client machine.

### Targets Built
1. **`DefFileGenerator-GUI.exe`**: Desktop GUI application (`DefFileGenerator/gui.py`), compiled with `console=False` (windowed, no command prompt popup). Bundles CustomTkinter json themes, Roboto fonts, and darkdetect.
2. **`deffilegen.exe`**: Command-line binary (`DefFileGenerator/main.py`), compiled with `console=True` for interactive PowerShell and automation pipeline usage.

### One-Click Local Compilation
```cmd
# Windows batch launcher (detects .venv or system Python):
build_exe.bat

# Python build orchestrator:
python build_exe.py --target all --zip
```
The orchestrator:
1. Cleans `build/` and `dist/` directories.
2. Invokes PyInstaller with `DefFileGenerator.spec`.
3. Runs automated smoke tests (`deffilegen.exe --version` and size checks).
4. Generates `dist/SHA256SUMS.txt` checksum file.
5. Packages a distribution archive: `dist/DefFileGenerator-v{version}-windows-x64.zip`.

---

[Back to README](README.md) · [Contributor checklist](docs/development.md) · [Security model](docs/security.md)
