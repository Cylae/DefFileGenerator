# DefFileGenerator

This toolset allows for extracting Modbus register information from manufacturer documentation (PDF, Excel, CSV, or XML) and generating WebdynSunPM definition files (CSV format). It handles address formatting, type validation, overlap detection, coefficient calculation, and address offsets.

## Key Features

*   **Robust Extraction**: Heuristic-based column detection for manufacturer documents.
*   **Secure XML Processing**: XXE-protected XML parsing via `defusedxml`.
*   **Advanced Address Logic**:
    *   Supports Decimal, Hex (0x prefix or h suffix), and Negative addresses.
    *   `address_offset`: Shift all register addresses by a specified value.
    *   Optimized overlap detection for large-scale register maps.
    *   **Strict Range Validation**: Ensures Modbus addresses are within 0-65535.
*   **Comprehensive Type Support**: Standardizes synonyms and supports endianness suffixes (e.g., `_B`, `_W`, `_WB`).
*   **Unified CLI**: Single entry point for extraction, generation, or end-to-end runs.
*   **Intelligent Action Defaulting**: Automatically assigns Read Only (4) for Input/Discrete registers and Read/Write (1) for Holding/Coils.

## Requirements

*   Python 3.x
*   Dependencies: `pdfplumber`, `openpyxl`, `pandas`, `lxml`, `defusedxml`, `reportlab`

Install all dependencies:
```bash
pip install pdfplumber openpyxl pandas lxml defusedxml reportlab
```

## Unified CLI Usage

The primary entry point is `DefFileGenerator/main.py`.

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

### 4. Validate Definition File
Check a generated WebdynSunPM CSV for formatting and Modbus range constraints.

```bash
python3 DefFileGenerator/main.py validate <definition_file>
```

---

## Input CSV Format

The simplified CSV (input for `generate`) uses these columns:

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

### Supported Types

*   **Numeric**: `U8-U64`, `I8-I64`, `F32`, `F64` (supports `_W`, `_B`, `_WB` suffixes).
*   **Convenience**: `STR<n>` (e.g., `STR20` for a 20-character string).
*   **Special**: `BITS`, `IP`, `IPV6`, `MAC`.

## Column Recognition Patterns

The tool searches for columns matching these patterns (case-insensitive):

| Target | Patterns |
| :--- | :--- |
| **Address** | register, address, addr, offset, reg |
| **Name** | name, description, parameter, variable, signal |
| **Type** | type, data type, format, datatype |
| **Unit** | unit, units |
| **Scale** | scale, factor, multiplier, ratio |
| **Action** | action, access |

## Validation & Performance

The tool is optimized for performance and strict resource constraints. It performs:
*   **Memory Efficiency**: O(1) memory overhead through generator-based stream processing.
*   **Input Resilience**: Advanced IO error isolation (fallback encoding and explicit type hints).
*   **Address Overlap Detection**: Dictionary-based O(N) check avoiding geometric performance drops.
*   **Security Validation**: Blocks external entity injection (XXE) in XML formats reliably.
