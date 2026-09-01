# WebdynSunPM DefFileGenerator & Documentation Parser

A python toolset and library for extracting Modbus register maps from manufacturer documentation (PDF, Excel, CSV, or XML) and generating validated, properly formatted WebdynSunPM definition CSV files.

It manages complex register transformations including hex/decimal address normalization, data type standardizations with endianness flags, automated tag generation, scaling factor / offset calculation, register address overlap detection, and CSV formula injection protection.

---

## Key Features

*   **Multi-Format Documentation Parser**: Extract register maps dynamically from PDF datasheets, Excel workbooks (all sheets or specific sheets), CSV files, or structured XML documents.
*   **Security & Safety**:
    *   XXE-protected XML parsing using `defusedxml`.
    *   CSV injection protection via `sanitize_csv_field` which escapes formula triggers (`=`, `@`, `+`, `-`) while preserving signed numbers (e.g. `-10.5`).
*   **Advanced Modbus Address Handling**:
    *   Supports Decimal, Hexadecimal (`0x` prefix or `h` suffix), negative addresses, and compound string/bit formats (`address_startbit_length`).
    *   Apply custom integer `address_offset` to shift all extracted register addresses.
    *   High-performance O(log N) binary search interval lookup for detecting register address overlaps across coil, discrete, holding, and input registers.
*   **Comprehensive Data Type Normalization**:
    *   Automatic type mapping (e.g., `uint16` -> `U16`, `float` -> `F32`, `int32` -> `I32`).
    *   Supports endianness suffixes (`_B`, `_W`, `_WB`).
    *   String register support (`STR<n>` or `STRING`) and bitfield definitions (`BITS`).
*   **High Performance & Streaming**: Low $O(1)$ memory footprint using generator-based row streaming pipelines capable of processing thousands of registers efficiently.
*   **Unified CLI & Programmatic API**: Use via simple python module calls or command-line interfaces (`main.py`, `doc_to_webdyn.py`, `generate_webdyn_def.py`).

---

## 🚀 Installation & Requirements

### Installation

```bash
pip install pdfplumber openpyxl defusedxml lxml
```

*(Optional dependencies: `pytest` for testing, `pandas` and `reportlab` for benchmark/stress test suites).*

---

## 💻 Usage & Entry Points

### 1. Programmatic API (`generate_webdyn_def.py`)

Integrate WebdynSunPM generation directly into Python applications using `generate_webdyn_definition`:

```python
from generate_webdyn_def import generate_webdyn_definition

success = generate_webdyn_definition(
    input_file="solar_inverter_map.xlsx",
    output_file="webdyn_definition.csv",
    manufacturer="Huawei",
    model="SUN2000-50KTL",
    protocol="modbusRTU",      # default: modbusRTU
    category="Inverter",        # default: Inverter
    address_offset=0,           # optional address shift
    strict_validation=True      # fail on address overlaps or format errors
)

if success:
    print("Definition file successfully generated and validated!")
```

Run built-in programmatic demo:
```bash
python3 generate_webdyn_def.py
```

---

### 2. Multi-Command CLI (`DefFileGenerator/main.py`)

The primary CLI provides four distinct sub-commands (`run`, `extract`, `generate`, `validate`).

#### End-to-End Extraction & Generation (`run`)
Extract registers from any documentation file and output a validated WebdynSunPM definition CSV in a single command:

```bash
python3 DefFileGenerator/main.py run input_doc.pdf \
    --manufacturer "SolarEdge" \
    --model "SE10K" \
    -o solaredge_definition.csv
```

#### Extract Registers to Intermediate CSV (`extract`)
```bash
python3 DefFileGenerator/main.py extract datasheet.xlsx --sheet "Holding Registers" -o intermediate_registers.csv
```

#### Generate Definition from Intermediate CSV (`generate`)
```bash
python3 DefFileGenerator/main.py generate intermediate_registers.csv \
    --manufacturer "SMA" \
    --model "STP-5000TL" \
    -o sma_definition.csv
```

#### Validate Definition CSV (`validate`)
```bash
python3 DefFileGenerator/main.py validate webdyn_definition.csv
```

---

### 3. Single-Step CLI (`doc_to_webdyn.py`)

A simplified command-line interface for direct conversion:

```bash
python3 doc_to_webdyn.py input_file.pdf \
    --manufacturer "Fronius" \
    --model "Symo-5.0" \
    -o fronius_definition.csv
```

---

## 🔍 Extraction Heuristics & Mapping

The extractor automatically detects columns using exact header matching and heuristic substring search across common naming conventions:

| Internal Field | Detected Header Patterns |
| :--- | :--- |
| **Address** | `address`, `addr`, `offset`, `register`, `reg` |
| **Name** | `name`, `description`, `parameter`, `variable`, `signal`, `signal name` |
| **Type** | `data type`, `datatype`, `type`, `format` |
| **Unit** | `unit`, `units` |
| **Action** | `action`, `access` |
| **Factor** | `scale`, `factor`, `multiplier`, `ratio` |
| **Offset** | `offset`, `bias`, `coefficient b` |
| **ScaleFactor** | `scalefactor`, `scale factor` |
| **Length** | `length`, `len`, `size`, `count`, `quantity` |
| **StartBit** | `startbit`, `bit offset`, `bit`, `start` |

### Column Mapping Customization

You can provide a custom JSON mapping file to override header auto-detection:

```json
{
  "Address": "Modbus_Addr",
  "Name": "Signal_Description",
  "Type": "Format_Code",
  "Unit": "Engineering_Unit"
}
```

Pass the mapping file via CLI: `--mapping custom_mapping.json`.

---

## 📊 Expected WebdynSunPM Output CSV Format

The output definition CSV conforms strictly to the WebdynSunPM header and row layout:

```csv
modbusRTU;Inverter;Huawei;SUN2000-50KTL;;;;;;;
1;3;40001;U16;;Active Power;active_power;1.000000;0.000000;W;4
2;3;40002;U16;;Grid Voltage;grid_voltage;0.100000;0.000000;V;4
```

---

## 🧪 Testing

Run the full unit test suite using Python's standard `unittest` framework:

```bash
PYTHONPATH=. python3 -m unittest discover -s DefFileGenerator/tests
```

Run specialized stress and torture test batteries:
```bash
PYTHONPATH=. python3 DefFileGenerator/tests/run_torture_battery.py
PYTHONPATH=. python3 DefFileGenerator/tests/stress_test_gen.py
PYTHONPATH=. python3 DefFileGenerator/tests/run_gigantic_battery.py
```
