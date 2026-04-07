# WebdynSunPM Documentation Parser

This tool automatically extracts Modbus register information from manufacturer documentation files (PDF, Excel, CSV, XML) and generates WebdynSunPM definition files.

## Features

- **Automated Extraction**: Finds register tables in PDF, Excel, CSV, and XML files.
- **Heuristic Mapping**: Automatically identifies columns like Address, Name, Type, Unit, and Scale using common naming conventions.
- **Data Normalization**:
  - Converts hex addresses (e.g., `0x9C40`) to decimal.
  - Maps manufacturer-specific data types (e.g., `uint16`, `float32`) to Webdyn types (`U16`, `F32`).
  - Generates unique tags from register names (collapsed underscores and stripped).
- **Scaling Support**: Handles scaling factors, multipliers, and power-of-10 scaling.
- **Compound Addresses**: Supports `Address_StartBit_Length` and `Address_Length` formats automatically.
- **Address Offsets**: Shift all addresses by a user-defined numeric value.

## Installation

```bash
pip install pandas openpyxl pdfplumber lxml defusedxml
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
- `--pages`: Specific PDF pages to process (comma-separated integers).
- `--mapping`: Custom JSON mapping file for column identification.
- `--address-offset`: Integer value to shift all register addresses.
- `--forced-write`: Forced write configuration string for the header.
- `-v`, `--verbose`: Show detailed processing information.

## How It Works

### Column Recognition

The tool searches for columns matching these patterns (case-insensitive):

| Target | Patterns |
| :--- | :--- |
| **Address** | register, address, addr, register, reg |
| **Name** | name, description, parameter, variable, signal, signal name |
| **Type** | type, data type, format, datatype |
| **Unit** | unit, units |
| **Scale/Factor**| scale, factor, multiplier, ratio |
| **Offset** | offset, bias, coefficient b |
| **Action** | action, access |
| **Length** | length, len, size, count, quantity |
| **StartBit** | startbit, bit offset, bit |

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
| string | STRING |
| bits | BITS |

### Normalization Logic

- **Addresses**: Removes commas, extracts numbers, and converts hex to decimal. Supports compound formats for `BITS` and `STRING`.
- **Tags**: Replaces non-alphanumeric characters with underscores, collapses multiple underscores, and strips them from the ends. Ensures uniqueness within the definition file.
- **Scaling**: Calculates `CoefA` using `Factor * 10^ScaleFactor`. Supports fractions like `1/10`.

## Integration

This tool uses the core `Extractor` and `Generator` logic from `DefFileGenerator/` to ensure that the generated files follow all WebdynSunPM rules, including address overlap detection, type validation, and resource management via context managers.
