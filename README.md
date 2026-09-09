# WebdynSunPM DefFileGenerator & Documentation Parser

A production-grade Python library and CLI toolset for extracting Modbus register maps from heterogeneous manufacturer documentation formats (**PDF, Excel, CSV, XML**) and generating validated, standardized **WebdynSunPM definition CSV files**.

---

## 🏗️ System Architecture & Workflow

The system is designed around a decoupled, generator-based streaming pipeline:

```text
[ Document Source ] (PDF / Excel / CSV / XML)
         │
         ▼
[ Extractor Pipeline ] ──► (Lazy Row Generators & Two-Pass Heuristic Mapping)
         │
         ▼
[ Generator Core ]     ──► (Address Normalization, Type Conversion, Offset Shift)
         │
         ▼
[ Validation Layer ]   ──► (O(log N) Bisect Overlap Check, Tag Uniqueness, Range Check)
         │
         ▼
[ Output Writer ]      ──► (CSV Injection Sanitization, Webdyn Header & Line Formatting)
```

### 1. Extractor Layer (`DefFileGenerator/extractor.py`)
- **PDF Extraction**: Employs `pdfplumber` to extract structured visual tables page-by-page. Handles newline collapsing, merged cell whitespace normalization, and selective page range parsing.
- **Excel Parsing**: Uses `openpyxl` in `read_only=True` mode with `data_only=True` to stream cell values without evaluating formula trees or consuming memory for unused formatting.
- **CSV Parsing**: Uses `csv.Sniffer` to detect delimiters (`,`, `;`, `\t`) and character encodings (`utf-8-sig`, `utf-16` BOM), handling multi-dialect inputs.
- **XML Parsing**: Uses `defusedxml.ElementTree` to parse hierarchical register maps while enforcing strict protections against DTDs, XXE attacks, and entity expansion bombs.

### 2. Two-Pass Heuristic Column Mapping
The extractor dynamically identifies target metadata fields (`Address`, `Name`, `Type`, `Unit`, `Action`, `Factor`, `Gain`, `Offset`, `ScaleFactor`, `Length`, `StartBit`) from varied manufacturer headers using a three-tier algorithm:
1. **Target-Name Exact Match**: Case-insensitive exact name matching.
2. **Synonym Pattern Matching**: Matching against pre-compiled synonym dictionaries (e.g., `reg type`, `modbus type` -> `RegisterType`).
3. **Substring Partial Matching**: Fallback fuzzy match for complex headers (e.g. `Register Start Address` -> `Address`).

### 3. Modbus Address Normalization & Overlap Detection
- **Address Formatting**: Converts hex strings (`0x1000`, `1000h`), negative hex/decimal values, and compound register formats (`address_startbit_length` for bitfields; `address_bytes` for strings).
- **O(log N) Interval Bisect Lookup**: Overlap verification uses `bisect.bisect_left` on sorted interval lists `(start_addr, end_addr)`. Registers of width 1 (U16/I16/Coil), 2 (U32/I32/F32/IP), 4 (U64/I64/F64), 3 (MAC), or 8 (IPV6) are checked in logarithmic time against previously mapped registers, allowing real-time validation of large scale files (5000+ registers) without $O(N^2)$ bottlenecks.
- **Address Offset Application**: Applies an integer shift (`address_offset`) across all addresses while validating that resulting addresses remain within valid Modbus bounds ($0$ to $65535$).

### 4. Data Type Normalization & Coefficient Calculation
- **Types Supported**: `U8`, `I8`, `U16`, `I16`, `U32`, `I32`, `U64`, `I64`, `F32`, `F64`, `STRING`, `BITS`, `IP`, `IPV6`, `MAC`.
- **Endianness Flags**: Endianness hints in input strings (`swap`, `big endian`, `word`) automatically yield normalized suffixes (`_WB`, `_B`, `_W`).
- **Coefficients (`CoefA` & `CoefB`)**: Linear scaling parameters are calculated as:
  $$\text{CoefA} = \text{Factor} \times 10^{\text{ScaleFactor}}$$
  $$\text{CoefB} = \text{Offset}$$
  Both values are formatted cleanly to 6 decimal places.

---

## 🔒 Security Specifications

- **CSV Formula Injection Mitigation**: `Generator.sanitize_csv_field` prevents spreadsheet macro execution. Any field starting with formula triggers (`=`, `+`, `-`, `@`, `|`, `%`), fullwidth variants (`\uff1d`, `\uff0b`, `\uff0d`, `\uff20`), or leading control characters (`\t`, `\r`, `\n`, NBSP) has a single apostrophe (`'`) prepended. Valid numeric literals (`-10.5`, `+25`, `1.5e3`) are preserved as numbers.
- **Control Character Filtering**: Non-printable characters (such as embedded null bytes `\x00`) are stripped before writing output CSVs.
- **XXE & Bomb Resistance**: XML parsing strictly enforces `defusedxml` constraints (`DTDForbidden`, `EntitiesForbidden`, `ExternalReferenceForbidden`).

---

## 💻 Usage & CLI Reference

### 1. Primary CLI (`DefFileGenerator/main.py` / `deffilegen`)

#### Subcommands:
- `run`: Extracts registers from documentation and generates a validated definition CSV in a single pass.
- `extract`: Extracts raw registers into an intermediate normalized CSV.
- `generate`: Converts an intermediate register CSV into a WebdynSunPM definition CSV.
- `validate`: Validates an existing definition file for Tag uniqueness, Type validity, and Address overlaps.

#### Example Usage:
```bash
# Full conversion run
deffilegen run datasheet.pdf --manufacturer "Huawei" --model "SUN2000" -o webdyn_huawei.csv

# Extraction only
deffilegen extract inverter_map.xlsx --sheet "Registers" -o extracted.csv

# Validation
deffilegen validate webdyn_huawei.csv
```

### 2. Direct Programmatic API (`generate_webdyn_def.py`)

```python
from generate_webdyn_def import generate_webdyn_definition

success = generate_webdyn_definition(
    input_file="registers.xlsx",
    output_file="output_definition.csv",
    manufacturer="SMA",
    model="STP5000",
    protocol="modbusRTU",
    category="Inverter",
    address_offset=0,
    strict_validation=True,
)
```

---

## 🧪 Quality Assurance & Test Suite

The repository maintains 100% test pass rate across:
- **Unit & Integration Suite**: `pytest` (463 tests passing)
- **Static Type Checking**: `mypy DefFileGenerator generate_webdyn_def.py doc_to_webdyn.py`
- **Linting & Formatting**: `ruff check .` and `ruff format --check .`
- **Stress & Torture Batteries**:
  - `run_torture_battery.py`: Tests ambiguous column resolution, string length shifts, and extreme edge cases.
  - `run_gigantic_battery.py`: Benchmark execution across 5,000+ registers, memory streaming checks, and type normalization call volumes.
