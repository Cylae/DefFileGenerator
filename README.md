# DefFileGenerator & WebdynSunPM Documentation Parser

This toolset allows for automatically extracting Modbus register information from manufacturer documentation (PDF, Excel, CSV, or XML) and generating WebdynSunPM definition files (CSV format). It handles address formatting, type validation, overlap detection, coefficient calculation, and address offsets.

## Quick Start Guide

### What This Tool Does

This tool **automatically extracts** register information from manufacturer documentation files and generates WebdynSunPM definition files.

Simply provide a PDF, Excel, CSV, or XML file from the manufacturer, and it will:
1. Find the register tables
2. Extract addresses, names, data types, units, etc.
3. Generate a ready-to-use WebdynSunPM definition file

### Installation

```bash
# Install required dependencies
pip install pandas openpyxl pdfplumber defusedxml lxml
```

### Basic Usage

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

## Key Features

*   **Robust Extraction**: Heuristic-based column detection for manufacturer documents (PDF, Excel, CSV, XML).
*   **Secure XML Processing**: XXE-protected XML parsing via `defusedxml`.
*   **Advanced Address Logic**:
    *   Supports Decimal, Hex (0x prefix or h suffix), and Negative addresses.
    *   `address_offset`: Shift all register addresses by a specified value.
    *   Optimized overlap detection for large-scale register maps.
    *   Modbus range validation (0-65535).
*   **Comprehensive Type Support**: Standardizes synonyms and supports endianness suffixes (e.g., `_B`, `_W`, `_WB`).
*   **Intelligent Action Defaulting**: Automatically assigns Read-Only or Read/Write actions based on register type.
*   **Unified CLI**: Single entry point for extraction, generation, validation, or end-to-end runs.

## Unified CLI Usage

The primary entry point is `DefFileGenerator/main.py`. Use `PYTHONPATH=. python3 DefFileGenerator/main.py` if running from the root.

### 1. Extract registers from documentation
Extract tables from PDF, Excel, CSV, or XML into a simplified CSV format.

```bash
python3 DefFileGenerator/main.py extract <source_file> -o <output_csv> [options]
```
*   `--mapping <json_file>`: (Optional) JSON file to map manufacturer columns.
*   `--sheet <name>`: (Excel only) Specific sheet name.
*   `--pages <list>`: (PDF only) Comma-separated list of pages.

### 2. Generate definition from CSV
Convert a simplified CSV into a WebdynSunPM definition file.

```bash
python3 DefFileGenerator/main.py generate <input_csv> --manufacturer <Name> --model <Model> -o <output_def_csv> [options]
```
*   `--address-offset <int>`: Shift addresses (default 0).
*   `--template`: Generate a sample simplified CSV template.

### 3. End-to-End Run
Extract and generate the definition file in a single step.

```bash
python3 DefFileGenerator/main.py run <source_file> --manufacturer <Name> --model <Model> -o <output_def_csv> [options]
```

### 4. Validate a definition file
Validate a generated WebdynSunPM definition file for errors.

```bash
python3 DefFileGenerator/main.py validate <definition_csv>
```

---

## How It Works

### Step 1: Column Recognition
The tool searches for columns matching these patterns (case-insensitive):

| Target | Patterns |
| :--- | :--- |
| **Address** | register, address, addr, offset, reg |
| **Name** | name, description, parameter, variable, signal |
| **Type** | type, data type, format, datatype |
| **Unit** | unit, units |
| **Scale** | scale, factor, multiplier, ratio |
| **Action** | action, access |

### Step 2: Data Normalization
- **Addresses**: Removes commas, extracts numbers, and converts hex to decimal.
- **Data Types**: Maps manufacturer-specific types (e.g., `uint16`, `float32`) to Webdyn types (`U16`, `F32`).
- **Tags**: Lowercases and replaces non-alphanumeric characters with underscores. Ensures uniqueness.
- **Scaling**: Calculates `CoefA` and `CoefB`. Supports fractions like `1/10`.

### Step 3: WebdynSunPM Definition Generation
Outputs a properly formatted CSV file with:
- Correct header (protocol, category, manufacturer, model).
- Indexed entries.
- Validated addresses and types.
- Calculated coefficients.

---

## Troubleshooting

### Problem: No registers extracted
**Solution:**
1. Check if your file has clearly labeled columns.
2. Run with `-v` (verbose) to see detailed logs.
3. Ensure PDF tables are text-based, not scanned images.

### Problem: Wrong data types
**Solution:**
- Add a "Type" or "Data Type" column to your source file.
- The tool will default to `U16` if it cannot guess.

### Problem: Incorrect addresses
**Solution:**
- Verify addresses are in the identified "Address" column.
- The tool handles hex (`0x9C40`) and decimal automatically.

---

## Validation & Performance

*   **Memory Efficiency**: O(1) memory overhead through generator-based stream processing.
*   **Input Resilience**: Advanced IO error isolation and fallback encoding (UTF-8, UTF-16).
*   **Address Overlap Detection**: Dictionary-based O(N) check.
*   **Security**: XXE protection for XML formats.
