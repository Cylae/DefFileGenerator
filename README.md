# DefFileGenerator

This toolset allows for extracting Modbus register information from manufacturer documentation (PDF or Excel) and generating WebdynSunPM definition files (CSV format). It handles address formatting, type validation, overlap detection, and coefficient calculation.

## Requirements

*   Python 3.x
*   Dependencies: `pdfplumber`, `openpyxl`

Install dependencies:
```bash
pip install pdfplumber openpyxl
```

## Unified CLI Usage

The primary entry point is `DefFileGenerator/main.py`, which provides three main subcommands.

### 1. Extract registers from documentation
Extract tables from a PDF or Excel file into a simplified CSV format.

```bash
python3 DefFileGenerator/main.py extract <source_file> -o <output_csv> [options]
```
*   `--mapping <json_file>`: (Optional) JSON file to map manufacturer columns to standard columns.
*   `--sheet <name>`: (Excel only) Specific sheet name to extract from.
*   `--pages <list>`: (PDF only) Comma-separated list of pages (e.g., `1,2,5`).

### 2. Generate definition from CSV
Convert a simplified CSV into a WebdynSunPM definition file.

```bash
python3 DefFileGenerator/main.py generate <input_csv> --manufacturer <Name> --model <Model> -o <output_def_csv> [options]
```

### 3. End-to-End Run
Extract and generate the definition file in a single step.

```bash
python3 DefFileGenerator/main.py run <source_file> --manufacturer <Name> --model <Model> -o <output_def_csv> [options]
```

---

## Direct Generator Usage

You can also run the generator script directly if you already have a correctly formatted simplified CSV.

```bash
python3 DefFileGenerator/def_gen.py <input_file> --manufacturer <Manufacturer> --model <Model> [options]
```

### Examples

**End-to-end extraction and generation:**
```bash
python3 DefFileGenerator/main.py run manual.pdf --manufacturer "MyCompany" --model "InverterX" -o definition.csv
```

**Generate a template input file:**
```bash
python3 DefFileGenerator/def_gen.py --template -o template.csv
```

## Input CSV Format

If you are creating the simplified CSV manually or via extraction mapping, it should have the following columns:

| Column | Description | Example |
| :--- | :--- | :--- |
| `Name` | Variable name (Required). | `Voltage L1` |
| `Tag` | Unique tag (auto-generated if empty). | `voltage_l1` |
| `RegisterType` | `Holding Register`, `Input Register`, etc. | `Holding Register` |
| `Address` | Register address (supports Dec, Hex like `0x10`, or `Addr_Len` for strings). | `30001` |
| `Type` | Data type (e.g., `U16`, `F32`, `STR20`). | `U16` |
| `Factor` | Multiplier factor (default 1). | `0.1` |
| `Offset` | Offset value (default 0). | `0` |
| `Unit` | Unit of measurement. | `V` |
| `ScaleFactor` | Power of 10 scaling ($CoefA = Factor \times 10^{ScaleFactor}$). | `-1` |

### Supported Types

*   **Integers**: `U8`, `U16`, `U32`, `U64`, `I8`, `I16`, `I32`, `I64` (supports `_W`, `_B`, `_WB` suffixes).
*   **Floats**: `F32`, `F64`
*   **Strings**: `STRING` (requires `Address_Length`), `STR<n>` (e.g., `STR20`).
*   **Bits**: `BITS` (requires `Address_StartBit_NumBits`).
*   **Network**: `IP`, `IPV6`, `MAC`.

## Validation & Checks

The tool performs the following checks:
*   **Type Validation**: Ensures the specified type is supported.
*   **Address Validation**: Checks if the address format matches the type.
*   **Duplicate Detection**: Warns if multiple rows have the same `Name` or `Tag`.
*   **Overlap Detection**: Warns if register addresses overlap.
