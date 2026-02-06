# DefFileGenerator

This tool generates WebdynSunPM Modbus definition files (CSV format) from a simplified CSV input file. It handles address formatting, type validation, overlap detection, and coefficient calculation.

## Requirements

*   Python 3.x

## Usage

Run the script from the command line:

```bash
python3 DefFileGenerator/def_gen.py <input_file> --manufacturer <Manufacturer> --model <Model> [options]
```

### Arguments

*   `input_file`: Path to the simplified CSV input file.
*   `--manufacturer`: Manufacturer name (Required).
*   `--model`: Model name (Required).
*   `-o`, `--output`: Path to the output CSV file. If not specified, prints to stdout.
*   `--protocol`: Protocol name (Default: `modbusRTU`).
*   `--category`: Device category (Default: `Inverter`).
*   `--forced-write`: Forced write code (Default: empty).
*   `--template`: Generate a template input CSV file to the specified output path (or stdout).

### Examples

**Generate a definition file:**

```bash
python3 DefFileGenerator/def_gen.py input.csv --manufacturer "MyCompany" --model "SolarInverter" -o definition.csv
```

**Generate a template input file:**

```bash
python3 DefFileGenerator/def_gen.py --template -o template.csv
```

## Input CSV Format

The input CSV should have the following columns (headers are case-insensitive):

| Column | Description | Example |
| :--- | :--- | :--- |
| `Name` | Variable name (Required). | `Voltage L1` |
| `Tag` | Unique tag for the variable. | `voltage_l1` |
| `RegisterType` | Modbus register type (`Holding Register`, `Input Register`, `Coil`, `Discrete Input`). | `Holding Register` |
| `Address` | Register address. For `STRING` use `Addr_Len`. For `BITS` use `Addr_Start_Num`. | `30001` or `30010_20` |
| `Type` | Data type (see Supported Types). | `U16`, `F32` |
| `Factor` | Multiplier factor (default 1). | `0.1` |
| `Offset` | Offset value (default 0). | `0` |
| `Unit` | Unit of measurement. | `V` |
| `Action` | Action code (0, 1, 2, 4, 6, 7, 8, 9). Default is 1. | `4` |
| `ScaleFactor` | Power of 10 scaling for CoefA calculation ($CoefA = Factor \times 10^{ScaleFactor}$). | `-1` |

### Supported Types

*   **Integers**: `U8`, `U16`, `U32`, `U64`, `I8`, `I16`, `I32`, `I64`
    *   Suffixes `_W`, `_B`, `_WB` are supported (e.g., `U16_W`).
*   **Floats**: `F32`, `F64`
*   **Strings**: `STRING` (requires `Address_Length`), `STR<n>` (e.g., `STR20` - automatically sets length to 20).
*   **Bits**: `BITS` (requires `Address_StartBit_NumBits`).
*   **Network**: `IP`, `IPV6`, `MAC`.

## Validation & Checks

The script performs the following checks:
*   **Type Validation**: Ensures the specified type is supported.
*   **Address Validation**: Checks if the address format matches the type.
*   **Duplicate Detection**: Warns if multiple rows have the same `Name` or `Tag`.
*   **Overlap Detection**: Warns if register addresses overlap (except for `BITS` sharing the same register).

## Output Format

The output is a semicolon-delimited CSV file compatible with WebdynSunPM, containing a header row with device metadata followed by data rows.
