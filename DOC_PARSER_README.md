# WebdynSunPM Documentation Parser

This tool automatically extracts Modbus register information from manufacturer documentation files (PDF, Excel, CSV, XML) and generates WebdynSunPM definition files.

## Features

- **Automated Extraction**: Finds register tables in PDF, Excel, CSV, and XML files.
- **Heuristic Mapping**: Automatically identifies columns like Address, Name, Type, Unit, and Scale using common naming conventions.
- **Data Normalization**:
  - Converts hex addresses (e.g., `0x9C40`) to decimal.
  - Maps manufacturer-specific data types (e.g., `uint16`, `float32`) to Webdyn types (`U16`, `F32`).
  - Generates unique tags from register names.
- **Scaling Support**: Handles scaling factors and multipliers.

## Installation

```bash
pip install pandas openpyxl pdfplumber lxml defusedxml reportlab
```

## Usage

```bash
python doc_to_webdyn.py INPUT_FILE --manufacturer MFG --model MODEL [OPTIONS]
```

### Arguments

- `INPUT_FILE`: Path to the manufacturer documentation (PDF, XLSX, XLS, CSV, XML).
- `--manufacturer`: Manufacturer name (Required).
- `--model`: Model name (Required).
- `-o`, `--output`: Output filename (default: `mfg_model_definition.csv`).
- `--protocol`: Protocol name (default: `modbusRTU`).
- `--category`: Device category (default: `Inverter`).
- `--sheet`: Specific Excel sheet name to process (processes all if omitted).
- `--pages`: Comma-separated list of PDF pages to process.
- `--mapping`: JSON file for custom column mapping.
- `--address-offset`: Value to shift all register addresses (default: 0).
- `--forced-write`: Global forced-write setting.
- `-v`, `--verbose`: Show detailed processing information.

## How It Works

### Column Recognition

The tool searches for columns matching these patterns (case-insensitive):

| Target | Patterns |
| :--- | :--- |
| **Address** | register, address, addr, offset, reg |
| **Name** | name, description, parameter, variable, signal |
| **Type** | type, data type, format, datatype |
| **Unit** | unit, units |
| **Scale** | scale, factor, multiplier, ratio |
| **Action** | action, access |

### Data Type Mapping

Common types are automatically mapped:

| Manufacturer Type | Webdyn Type |
| :--- | :--- |
| uint16, u16 | U16 |
| int16, i16 | I16 |
| uint32, u32 | U32 |
| int32, i32 | I32 |
| float, f32, float32 | F32 |
| double, f64, float64 | F64 |

### Normalization Logic

- **Addresses**: Removes commas, extracts numbers, and converts hex to decimal.
- **Tags**: Lowercases and replaces non-alphanumeric characters with underscores. Ensures uniqueness by adding numeric suffixes if necessary.
- **Scaling**: If a scale column is found, it is used as `CoefA`. Supports fractions like `1/10`.

## Integration

This tool uses the core `Generator` logic from `DefFileGenerator/def_gen.py` to ensure that the generated files follow all WebdynSunPM rules, including address overlap detection and type validation.
