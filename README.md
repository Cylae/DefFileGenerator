# WebdynSunPM DefFileGenerator & Documentation Parser

This toolset allows for extracting Modbus register information from manufacturer documentation (PDF, Excel, CSV, or XML) and automatically generating WebdynSunPM definition files (CSV format). It handles address formatting, type validation, overlap detection, coefficient calculation, and address offsets.

## Key Features

*   **Robust Extraction**: Heuristic-based column detection for manufacturer documents (finds Address, Name, Type, Unit, Scale, etc.).
*   **Secure XML Processing**: XXE-protected XML parsing via `defusedxml`.
*   **Advanced Address Logic**:
    *   Supports Decimal, Hex (0x prefix or h suffix), and Negative addresses.
    *   `address_offset`: Shift all register addresses by a specified value.
    *   Optimized overlap detection for large-scale register maps.
*   **Comprehensive Type Support**: Standardizes synonyms and supports endianness suffixes (e.g., `_B`, `_W`, `_WB`). Automatically maps types (e.g., `uint16` -> `U16`, `float` -> `F32`).
*   **Unified CLI**: Single entry point for extraction, generation, or end-to-end runs.
*   **High Performance**: O(1) memory overhead through generator-based stream processing, allowing handling of 5,000+ registers seamlessly.

---

## 🚀 A-Z Quick Start Guide

### 1. Installation

First, install the required dependencies:

```bash
pip install pdfplumber openpyxl defusedxml lxml
```
*(Optional: `pandas` for stress testing, `reportlab` for PDF generation in tests)*

### 2. Basic End-to-End Usage

The primary entry point is `DefFileGenerator/main.py`. The simplest way to use the tool is the `run` command, which extracts from documentation and generates the final definition file in one go.

```bash
python3 DefFileGenerator/main.py run INPUT_FILE --manufacturer "MFG" --model "MODEL" -o output.csv
```

**Examples with Real Files:**

**From PDF Documentation:**
```bash
python3 DefFileGenerator/main.py run manufacturer_datasheet.pdf \
    --manufacturer "Huawei" \
    --model "SUN2000-5KTL" \
    -o huawei_definition.csv
```

**From Excel Register Map:**
```bash
# Process all sheets
python3 DefFileGenerator/main.py run register_map.xlsx \
    --manufacturer "SolarEdge" \
    --model "SE5000H" \
    -o solaredge_definition.csv

# Process specific sheet
python3 DefFileGenerator/main.py run register_map.xlsx \
    --sheet "Holding Registers" \
    --manufacturer "SMA" \
    --model "STP-5000TL" \
    -o sma_definition.csv
```

**From CSV Export:**
```bash
python3 DefFileGenerator/main.py run registers.csv \
    --manufacturer "Fronius" \
    --model "Symo-5.0" \
    -o fronius_definition.csv
```

### 3. Step-by-Step CLI Commands

The CLI also allows splitting the process into extraction and generation:

#### Step 3A: Extract registers from documentation
Extract tables into a simplified intermediate CSV format.

```bash
python3 DefFileGenerator/main.py extract <source_file> -o <output_csv> [options]
```
*   `--mapping <json_file>`: (Optional) JSON file to explicitly map manufacturer columns.
*   `--sheet <name>`: (Excel only) Specific sheet name to process.
*   `--pages <list>`: (PDF only) Comma-separated list of pages to parse (e.g. `1,2,5`).
*   `--address-offset <int>`: Shift all addresses by a specified integer.

#### Step 3B: Generate definition from intermediate CSV
Convert the simplified intermediate CSV into a WebdynSunPM definition file.

```bash
python3 DefFileGenerator/main.py generate <input_csv> --manufacturer <Name> --model <Model> -o <output_def_csv> [options]
```
*   `--protocol <PROTO>`: Protocol name (default: `modbusRTU`).
*   `--category <CAT>`: Device category (default: `Inverter`).

### 4. How It Works (Extraction Heuristics)

The tool searches for columns matching these patterns (case-insensitive):

| Target | Patterns |
| :--- | :--- |
| **Address** | register, address, addr, offset, reg |
| **Name** | name, description, parameter, variable, signal |
| **Type** | type, data type, format, datatype |
| **Unit** | unit, units |
| **Scale** | scale, factor, multiplier, ratio |
| **Action** | action, access |

**Data Type Mapping:**
Common types are automatically mapped:
* `uint16`, `u16` -> `U16`
* `int16`, `i16` -> `I16`
* `uint32`, `u32` -> `U32`
* `float`, `f32`, `float32` -> `F32`

**Normalization Logic:**
* **Addresses**: Removes commas, extracts numbers, and converts hex (e.g., `0x9C40`) to decimal.
* **Tags**: Lowercases and replaces non-alphanumeric characters with underscores. Ensures uniqueness.
* **Scaling**: If a scale column is found, it is used as `CoefA`. Supports fractions like `1/10`.

### 5. Expected Output Format

The output is a properly formatted WebdynSunPM definition CSV:

```csv
modbusRTU;Inverter;Huawei;SUN2000-5KTL;;;;;;;
1;3;40001;U16;;Active Power;active_power;1.000000;0.000000;W;4
2;3;40002;U16;;Voltage;voltage;0.100000;0.000000;V;4
...
```

### 6. Troubleshooting

* **Problem: No registers extracted**
  * Check if your file has clearly labeled columns (e.g., "Address", "Name", "Type").
  * Run with `-v` (verbose) to see detailed processing information:
    ```bash
    python3 DefFileGenerator/main.py run yourfile.pdf --manufacturer "X" --model "Y" -v
    ```
  * Make sure tables in PDF are text-based (not scanned images).
* **Problem: Wrong data types**
  * Add a "Type" or "Data Type" column to your source file if missing. The tool guesses where possible.
* **Problem: Incorrect addresses**
  * Check if addresses are in the correctly matched column. The tool handles hex and decimal automatically.
* **Problem: Missing units or scaling**
  * These are optional, but if they exist, ensure the column headers contain "Unit", "Scale", or "Factor".

---

## Input CSV Format (Intermediate Format)

If you manually create the input CSV for the `generate` command, use these columns:

| Target | Patterns |
| :--- | :--- |
| `Name` | Variable name (Required). |
| `Tag` | Unique tag (auto-generated if empty). |
| `RegisterType` | e.g., `Holding Register`, `Input Register`. |
| `Address` | Register address (Dec, Hex like `0x10`, or `Addr_Len` for strings). |
| `Type` | Data type (e.g., `U16`, `F32_WB`, `STR20`). |
| `Factor` | Multiplier factor (supports fractions like `1/10`). |
| `Offset` | Offset value (default 0). |
| `Unit` | Unit of measurement. |
| `ScaleFactor` | Power of 10 scaling ($CoefA = Factor \times 10^{{ScaleFactor}}$). |
