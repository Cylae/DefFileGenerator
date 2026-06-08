# WebdynSunPM Documentation Parser - Quick Start Guide

## What This Tool Does

This tool **automatically extracts** register information from manufacturer documentation files and generates WebdynSunPM definition files.

Simply provide a PDF, Excel, CSV, or XML file from the manufacturer, and it will:
1. Find the register tables
2. Extract addresses, names, data types, units, etc.
3. Generate a ready-to-use WebdynSunPM definition file

## Installation

```bash
# Core execution dependencies
pip install openpyxl pdfplumber defusedxml lxml reportlab

# Optional dependencies (for testing/benchmarking)
pip install coverage pandas
```

## Basic Usage

### From PDF Documentation
```bash
python doc_to_webdyn.py manufacturer_datasheet.pdf \
    --manufacturer "Huawei" \
    --model "SUN2000-5KTL" \
    -o huawei_definition.csv
```

### From Excel Register Map
```bash
python doc_to_webdyn.py register_map.xlsx \
    --manufacturer "SolarEdge" \
    --model "SE5000H" \
    -o solaredge_definition.csv
```

### From CSV File
```bash
python doc_to_webdyn.py registers.csv \
    --manufacturer "Fronius" \
    --model "Symo-5.0" \
    -o fronius_definition.csv
```

## Advanced Features

### Generate Template
Create a sample Webdyn-compatible CSV register map to see the expected format:
```bash
python doc_to_webdyn.py --template -o template.csv
```

### Validate Definition File
Check an existing Webdyn definition file for errors (duplicates, overlaps):
```bash
python3 DefFileGenerator/main.py validate definition_file.csv
```

## How It Works

### Step 1: The tool looks for register information in your file

It searches for columns with names like:
- **Address**: register, address, addr, offset
- **Name**: name, description, parameter
- **Type**: type, data type, format
- **Unit**: unit, units
- **Scale**: scale, factor, multiplier
- etc.

### Step 2: It converts the data

- Normalizes addresses (handles hex like 0x9C40 or decimal like 40001)
- Converts data types (uint16 → U16, int32 → I32, float → F32, etc.)
- Generates unique tags from register names
- Calculates scaling coefficients (CoefA, CoefB)

### Step 3: Creates WebdynSunPM definition file

Outputs a properly formatted CSV file ready for WebdynSunPM:
```csv
modbusRTU;Inverter;Huawei;SUN2000-5KTL;;;;;;;
1;3;40001;U16;;Active Power;active_power;1.000000;0.000000;W;4
2;3;40002;U16;;Voltage;voltage;0.100000;0.000000;V;4
...
```

## Command-Line Options (doc_to_webdyn.py)

```bash
python doc_to_webdyn.py [INPUT_FILE] --manufacturer MFG --model MODEL [OPTIONS]
```

### Required Arguments
- `INPUT_FILE` - Your PDF, Excel, CSV, or XML file (optional if using --template)
- `--manufacturer MFG` - Manufacturer name
- `--model MODEL` - Model name

### Optional Arguments
- `-o OUTPUT` - Output filename (default: auto-generated)
- `--protocol PROTO` - Protocol name (default: modbusRTU)
- `--category CAT` - Device category (default: Inverter)
- `--sheet NAME` - Excel sheet name (processes all if not specified)
- `--pages PAGES` - PDF pages (e.g., "1,2,5")
- `--mapping FILE` - JSON mapping file for custom column detection
- `--address-offset N` - Shift all addresses by N
- `--forced-write STR` - Value for the forced write field
- `--template` - Generate a template file
- `-v, --verbose` - Show detailed processing information

## Troubleshooting

### Problem: No registers extracted
- Run with `-v` (verbose) to see detailed detection logs.
- Ensure tables in PDF are text-based, not scanned images.
- Check if your column headers are recognized (see DOC_PARSER_README.md).

### Problem: Address overlaps
- The tool detects if multiple registers occupy the same Modbus address.
- Use `DefFileGenerator/main.py validate` to identify specific overlap locations.

---

**You're ready to go! Just point the tool at your manufacturer documentation and it will do the rest.**
