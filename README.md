# DefFileGenerator & WebdynSunPM Documentation Parser

This toolset allows for automatically extracting Modbus register information from manufacturer documentation (PDF, Excel, CSV, or XML) and generating WebdynSunPM definition files (CSV format). It handles address formatting, type validation, overlap detection, coefficient calculation, and address offsets.

## Quick Start

### 1. Installation
```bash
# Install core and optional dependencies
pip install pandas openpyxl pdfplumber lxml defusedxml reportlab
```

### 2. Basic Usage (End-to-End)
```bash
python3 DefFileGenerator/main.py run datasheet.pdf --manufacturer "Huawei" --model "SUN2000" -o definition.csv
```

---

## Key Features

*   **Robust Extraction**: Heuristic-based column detection for manufacturer documents (PDF, Excel, CSV, XML).
*   **Secure XML Processing**: XXE-protected XML parsing via `defusedxml`.
*   **Advanced Address Logic**:
    *   Supports Decimal, Hex (0x prefix or h suffix), and Negative addresses.
    *   Address Range Validation (0-65535).
    *   `address_offset`: Shift all register addresses by a specified value.
    *   Optimized overlap detection for large-scale register maps.
*   **Comprehensive Type Support**: Standardizes synonyms and supports endianness suffixes (e.g., `_B`, `_W`, `_WB`).
*   **Intelligent Action Defaulting**: Automatically assigns Read-Only (4) or Read/Write (1) based on register type.
*   **Unified CLI**: Single entry point for extraction, generation, validation, or end-to-end runs.

## Unified CLI Usage

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
Extract tables from PDF, Excel, CSV, or XML into a simplified CSV format.
```bash
python3 DefFileGenerator/main.py extract <source_file> -o <output_csv> [options]
```
*   `--mapping <json_file>`: JSON file to map manufacturer columns.
*   `--sheet <name>`: (Excel only) Specific sheet name to process.
*   `--pages <list>`: (PDF only) Comma-separated list of pages (e.g., "1,2,5").

### 2. Generate definition from CSV
Convert a simplified CSV into a WebdynSunPM definition file.
```bash
python3 DefFileGenerator/main.py generate <input_csv> --manufacturer <Name> --model <Model> -o <output_def_csv> [options]
```
*   `--template`: Generate a sample input CSV.

### 3. End-to-End Run
Extract and generate the definition file in a single step.
```bash
python3 DefFileGenerator/main.py run <source_file> --manufacturer <Name> --model <Model> -o <output_def_csv> [options]
```

### 4. Validate Definition
Validate a generated definition file for errors or overlaps.
```bash
python3 DefFileGenerator/main.py validate <definition_csv>
```

---

## Input CSV Format (Simplified)

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
| `Action` | Action code (1=RW, 4=RO, etc. Defaults based on type). |
| `ScaleFactor` | Power of 10 scaling ($CoefA = Factor \times 10^{ScaleFactor}$). |

---

## Troubleshooting

## Column Recognition (Heuristics)

The tool automatically identifies columns like:
- **Address**: register, address, addr, offset
- **Name**: name, description, parameter, variable
- **Type**: type, data type, format
- **Unit**: unit, units
- **Scale**: scale, factor, multiplier

## Validation & Performance

The tool is optimized for performance and strict resource constraints:
*   **Memory Efficiency**: O(1) memory overhead through generator-based stream processing.
*   **Address Overlap Detection**: Dictionary-based O(N) check.
*   **Security**: XXE protection for XML.
