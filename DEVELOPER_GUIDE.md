# DefFileGenerator — Developer Architecture & Onboarding Guide

Welcome to the **DefFileGenerator** project! This guide is written for software engineers maintaining, debugging, or extending this codebase. It documents the domain concepts, system architecture, data structures, invariants, and standard development workflows.

---

## 1. Executive Summary & Purpose

`DefFileGenerator` automates the conversion of vendor-specific Modbus documentation into syntactically valid, validated **WebdynSunPM** definition files (`.csv`).

Manufacturer register maps vary wildly in format:
- Different file types (Adobe PDF, Microsoft Excel, CSV, XML).
- Inconsistent column headers (`"Register"`, `"Addr (dec)"`, `"Point"`, `"Grandeur"`, `"Data Format"`).
- Diverse data types (`"uint16"`, `"FLOAT32 Big Endian"`, `"Bitfield16"`, `"STR20"`).
- Scaling variations (`"Gain: 10"` divisor vs. `"Multiplier: 0.1"` vs. `"Scale Factor: -1"`).

`DefFileGenerator` provides a format-agnostic extraction and normalization pipeline accessible via:
1. **Command Line Interface (`deffilegen`)** — Primary automation tool for pipelines and scripting.
2. **REST API Backend (`web/app.py`)** — FastAPI server powering browser-based conversion.
3. **Windows 11 Desktop GUI (`DefFileGenerator/gui.py`)** — Modern CustomTkinter application with live register preview.

---

## 2. System Architecture & Data Flow

The codebase strictly enforces a layered architecture where the core domain logic has zero dependency on UI or web frameworks:

```
[ PDF / Excel / CSV / XML ]
            │
            ▼
┌────────────────────────────────────────────────────────┐
│  DefFileGenerator/extractor.py                         │
│  - Stream buffer ingestion (avoids OS file locks)      │
│  - Smart header detection & Modbus keyword scoring     │
│  - Multi-row header banner merging                     │
│  - Statistical column type inference fallback          │
│  - Two-pass heuristic column mapping                   │
└───────────────────────────┬────────────────────────────┘
                            │ Yields intermediate register dicts
                            ▼
┌────────────────────────────────────────────────────────┐
│  DefFileGenerator/def_gen.py                           │
│  - Data type normalization & endianness tagging        │
│  - Address boundary validation (0..65535)              │
│  - O(log N) bisect interval address overlap detection  │
│  - Linear coefficient computation (CoefA, CoefB)       │
│  - Unique variable tag synthesis                       │
│  - CSV injection neutralization                        │
│  - Atomic disk persistence with fsync                  │
└───────────────────────────┬────────────────────────────┘
                            │ Generates validated definition CSV
                            ▼
        ┌───────────────────┼───────────────────┐
        ▼                   ▼                   ▼
┌───────────────┐   ┌───────────────┐   ┌───────────────┐
│      CLI      │   │  FastAPI Web  │   │  Windows GUI  │
│    main.py    │   │  web/app.py   │   │    gui.py     │
└───────────────┘   └───────────────┘   └───────────────┘
```

---

## 3. WebdynSunPM Specification & Domain Invariants

WebdynSunPM data loggers communicate with Modbus RTU/TCP inverters, meters, and weather stations. The gateway expects a semicolon-delimited (`;`), UTF-8-sig or UTF-16 definition CSV file structured as follows:

### Line 1: Header Row
| Column Index | Field Name | Description | Example |
|---|---|---|---|
| 0 | `Protocol` | Communication protocol | `modbusRTU` or `modbusTCP` |
| 1 | `Category` | Equipment classification | `Inverter`, `Meter`, `Sensor` |
| 2 | `Manufacturer` | Equipment brand | `Huawei`, `ABB`, `SMA` |
| 3 | `Model` | Device model identifier | `SUN2000`, `STP5000` |
| 4 | `ForcedWrite` | Optional forced write command | Empty or command string |
| 5–10 | Empty | Remaining columns on line 1 are empty semicolons | `;;;;;;` |

### Lines 2+: Register Rows
| Column Index | Field Name | Type | Description |
|---|---|---|---|
| 0 | `Index` | Integer | 1-based sequential line counter (`1`, `2`, `3`, ...) |
| 1 | `Info1` | Integer | Modbus function code category: `1`=Coil, `2`=Discrete Input, `3`=Holding Register, `4`=Input Register |
| 2 | `Info2` | String | Modbus address (e.g. `40001`, `40001_20` for strings, `40001_0_16` for bits) |
| 3 | `Info3` | String | Normalized Webdyn type: `U16`, `I16`, `U32_WB`, `F32_WB`, `STRING`, `BITS`, etc. |
| 4 | `Info4` | String | Reserved for future firmware expansion (always empty string `""`) |
| 5 | `Name` | String | Human-readable variable description (e.g. `"Active Power"`) |
| 6 | `Tag` | String | Unique variable identifier, lowercase starting with a letter (`active_power`) |
| 7 | `CoefA` | Float | Multiplier factor formatted to 6 decimal places ($a \cdot 10^{\text{scale}}$) |
| 8 | `CoefB` | Float | Offset addition bias formatted to 6 decimal places |
| 9 | `Unit` | String | Engineering measurement unit (`V`, `A`, `W`, `kW`, `Hz`, `°C`) |
| 10 | `Action` | String | Access permission: `"4"`=Read-Only (RO), `"1"`=Write/Read-Write (RW) |

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
- **Binary Search**: Uses `bisect.bisect_left` on sorted interval lists for $O(\log N)$ lookup performance.
- **Bit-Slice Exception**: Multiple `BITS` entries inside the same 16-bit word (e.g. `30001_0_1` and `30001_1_1`) do **not** trigger a collision if their bit ranges $[start, start + length - 1]$ are disjoint.

### 4.2. CSV Formula Injection Defense (`sanitize_csv_field`)
Spreadsheet applications (Microsoft Excel, LibreOffice Calc) execute formulas if a cell begins with `=, +, -, @, |, %`, or certain tab/whitespace triggers.
- Any cell starting with these triggers has an apostrophe (`'`) prepended.
- Finite numeric literals (e.g. `-10`, `+5.4`) are recognized and preserved without apostrophes.
- Multiline newlines (`\r`, `\n`) are replaced with single spaces to prevent CSV row desynchronization.

### 4.3. Safe File Operations
When writing output CSV files:
1. Data is written to a temporary sibling file (`.{name}.{uuid}.tmp`) in the destination directory.
2. Changes are flushed and synced to disk via `os.fsync(fd)`.
3. The file is atomically replaced via `os.replace`, guaranteeing that interruptions or power losses never produce corrupted partial files.

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

---

## 6. Development & Quality Assurance Workflows

All code must pass static typing, linting, and automated tests before committing.

### Running Tests
```powershell
# Run entire test suite (590+ tests)
pytest

# Run tests fast without external equipment benchmark
pytest DefFileGenerator/tests/test_def_gen.py DefFileGenerator/tests/test_extractor.py
```

### Linting & Code Formatting
```powershell
# Check for lint violations
ruff check .

# Check formatting compliance
ruff format --check .

# Auto-format all code
ruff format .
```

### Static Type Checking
```powershell
mypy DefFileGenerator web
```

### Static Security Analysis
```powershell
bandit -r DefFileGenerator web
```

### Packaging & Build
```powershell
python -m build
```

---

## 7. Standalone Windows Executables (.exe)

The repository provides a complete, automated PyInstaller pipeline to compile self-contained 64-bit Windows executables requiring zero Python runtime on the client machine.

### Targets Built
1. **`DefFileGenerator-GUI.exe`**: Desktop GUI application (`DefFileGenerator/gui.py`), compiled with `console=False` (windowed, no command prompt popup). Bundles CustomTkinter json themes, Roboto fonts, and darkdetect.
2. **`deffilegen.exe`**: Command-line binary (`DefFileGenerator/main.py`), compiled with `console=True` for interactive PowerShell and automation pipeline usage.

### PyInstaller Specification (`DefFileGenerator.spec`)
The spec file manages:
- **Asset Collection**: Dynamically collects non-code assets and hidden imports for `customtkinter`, `openpyxl`, `pdfplumber`, `pypdfium2`, and `defusedxml` via `PyInstaller.utils.hooks.collect_all`.
- **Target Selection**: Respects the `BUILD_TARGET` environment variable (`all`, `gui`, or `cli`).
- **Binary Trimming**: Excludes unnecessary heavyweight scientific packages (`matplotlib`, `scipy`, `pandas`, `IPython`) to minimize binary footprint and startup latency.

### One-Click Local Compilation
Developers can build both executables locally on Windows:
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

### CI/CD Release Automation (`.github/workflows/build-exe.yml`)
- **Automated Workflow Artifacts**: On every push to `main`, GitHub Actions compiles the executables on `windows-latest`, validates CLI execution, and uploads the `.exe` and `.zip` artifacts (14-day retention).
- **Automated GitHub Releases**: On every tag matching `v*` (e.g., `git tag v0.2.2 && git push --tags`), the workflow automatically drafts and publishes a GitHub Release with the standalone `.exe` binaries, zip bundle, and SHA-256 checksums.
