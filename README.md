# DefFileGenerator & WebdynSunPM Documentation Parser

This toolset allows for automatically extracting Modbus register information from manufacturer documentation (PDF, Excel, CSV, or XML) and generating WebdynSunPM definition files (CSV format). It handles address formatting, type validation, overlap detection, coefficient calculation, and address offsets.

## Quick Start Guide

### What This Tool Does
Simply provide a PDF, Excel, CSV, or XML file from the manufacturer, and it will:
1. **Find** the register tables using heuristic-based column detection.
2. **Extract** addresses, names, data types, units, and scaling factors.
3. **Generate** a properly formatted WebdynSunPM definition file.

### Installation
```bash
# Install core and optional dependencies
pip install pandas openpyxl pdfplumber lxml defusedxml reportlab
```

### Basic Usage
The `doc_to_webdyn.py` script provides a simple interface for most users.

#### From PDF Documentation
```bash
python doc_to_webdyn.py manufacturer_datasheet.pdf \
    --manufacturer "Huawei" \
    --model "SUN2000-5KTL" \
    -o huawei_definition.csv
```

#### From Excel Register Map
```bash
python doc_to_webdyn.py register_map.xlsx \
    --manufacturer "SolarEdge" \
    --model "SE5000H" \
    -o solaredge_definition.csv
```

#### From CSV File
```bash
python doc_to_webdyn.py registers.csv \
    --manufacturer "Fronius" \
    --model "Symo-5.0" \
    -o fronius_definition.csv
```

---

## Unified CLI (Advanced Usage)

The primary entry point for granular control is `DefFileGenerator/main.py`.

### 1. Extract registers from documentation
Extract tables from source files into a simplified internal CSV format.
```bash
python3 DefFileGenerator/main.py extract <source_file> -o <output_csv> [options]
```
*   `--mapping <json_file>`: JSON file to map manufacturer columns.
*   `--sheet <name>`: (Excel only) Specific sheet name to process.
*   `--pages <list>`: (PDF only) Comma-separated list of pages (e.g., "1,2,5").

### 2. Generate definition from CSV
Convert a simplified CSV (manually created or extracted) into a WebdynSunPM definition file.
```bash
python3 DefFileGenerator/main.py generate <input_csv> --manufacturer <Name> --model <Model> -o <output_def_csv> [options]
```
*   `--address-offset <int>`: Shift all addresses by a specific value.

### 3. End-to-End Run
Perform extraction and generation in a single command.
```bash
python3 DefFileGenerator/main.py run <source_file> --manufacturer <Name> --model <Model> -o <output_def_csv> [options]
```

---

## Technical Specifications

### Column Recognition
The tool automatically identifies columns matching these patterns (case-insensitive):

| Target | Patterns |
| :--- | :--- |
| **Address** | register, address, addr, offset, reg |
| **Name** | name, description, parameter, variable, signal, signal name |
| **Type** | type, data type, format, datatype |
| **Unit** | unit, units |
| **Scale** | scale, factor, multiplier, ratio |
| **Action** | action, access |

### Data Type Mapping
Standard manufacturer types are mapped to WebdynSunPM types:

| Manufacturer Type | Webdyn Type |
| :--- | :--- |
| uint16, u16 | U16 |
| int16, i16 | I16 |
| uint32, u32 | U32 |
| int32, i32 | I32 |
| float, f32, float32 | F32 |
| double, f64, float64 | F64 |
| string, str | STR<n> |

### Features & Security
*   **Address Logic**: Supports Decimal, Hex (0x prefix), and Negative addresses.
*   **Overlap Detection**: Dictionary-based O(N) check for register collisions.
*   **Security**: XXE-protected XML parsing via `defusedxml`.
*   **Memory Efficiency**: O(1) memory overhead using generator-based stream processing.

---

## Input CSV Format (for `generate`)

The simplified CSV uses these columns:

| Column | Description |
| :--- | :--- |
| `Name` | Variable name (Required). |
| `Tag` | Unique tag (auto-generated if empty). |
| `RegisterType` | e.g., `Holding Register`, `Input Register`. |
| `Address` | Register address (Dec, Hex like `0x10`, or `Addr_Len` for strings). |
| `Type` | Data type (e.g., `U16`, `F32_WB`, `STR20`). |
| `Factor` | Multiplier factor (supports fractions like `1/10`). |
| `Offset` | Offset value (default 0). |
| `Unit` | Unit of measurement. |
| `ScaleFactor` | Power of 10 scaling ($CoefA = Factor \times 10^{ScaleFactor}$). |

---

## Troubleshooting

*   **No registers extracted**:
    1. Check if your file has clearly labeled columns.
    2. Run with `-v` (verbose) to see detection logs.
    3. Ensure PDF tables are text-based (not scanned images).
*   **Wrong data types**: Add a "Type" or "Data Type" column to your source file for explicit mapping.
*   **Incorrect addresses**: Verify the "Address" column. The tool handles 0x hex and decimal automatically.

## Verification

Run the full test suite to ensure everything is working correctly:
```bash
PYTHONPATH=. python3 -m unittest discover DefFileGenerator/tests
```
