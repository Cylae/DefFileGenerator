# WebdynSunPM Documentation Parser - Quick Start Guide

## What This Tool Does

This tool **automatically extracts** register information from manufacturer documentation files and generates WebdynSunPM definition files.

Simply provide a PDF, Excel, CSV, or XML file from the manufacturer, and it will:
1. Find the register tables
2. Extract addresses, names, data types, units, etc.
3. Generate a ready-to-use WebdynSunPM definition file

## Installation

```bash
# Install core dependencies
pip install openpyxl pdfplumber defusedxml lxml reportlab

# Optional: Install pandas and coverage for testing and benchmarking
pip install pandas coverage
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

### From XML Documentation
```bash
python doc_to_webdyn.py registers.xml \
    --manufacturer "InvertersInc" \
    --model "INV-500" \
    -o inverters_definition.csv
```

### Generating a Template
```bash
python doc_to_webdyn.py --template -o my_template.csv
```

## How It Works

### Step 1: The tool looks for register information in your file

It searches for columns with names like:
- **Address**: register, address, addr, offset, reg
- **Name**: name, description, parameter, variable, signal
- **Type**: type, data type, format, datatype
- **Unit**: unit, units
- **Scale**: scale, factor, multiplier, ratio
- **Action**: action, access

### Step 2: It converts the data

- Normalizes addresses (handles hex like 0x9C40, decimal like 40001, and shifts via `--address-offset`)
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

## Command-Line Options

```bash
python doc_to_webdyn.py [INPUT_FILE] --manufacturer MFG --model MODEL [OPTIONS]
```

### Arguments
- `INPUT_FILE` - Your PDF, Excel, CSV, or XML file (optional if using `--template`)
- `--manufacturer MFG` - Manufacturer name (e.g., "Huawei")
- `--model MODEL` - Model name (e.g., "SUN2000-5KTL")

### Advanced Options
- `-o, --output` - Output filename (default: auto-generated)
- `--protocol PROTO` - Protocol name (default: modbusRTU)
- `--category CAT` - Device category (default: Inverter)
- `--address-offset OFFSET` - Shift all addresses by this integer (default: 0)
- `--pages PAGES` - Specific PDF pages to process (e.g., "1,2,5-10")
- `--mapping JSON_FILE` - Custom column mapping file
- `--forced-write VAL` - Value for the forced write header field
- `--template` - Generate a sample CSV template instead of parsing a file
- `-v, --verbose` - Show detailed processing information

## Testing with Sample Files

Two sample files are included for testing:

### 1. CSV Sample
```bash
python doc_to_webdyn.py sample_register_map.csv \
    --manufacturer "TestMfg" \
    --model "TEST-1000" \
    -o test_csv_output.csv
```

### 2. Excel Sample
```bash
python doc_to_webdyn.py sample_inverter_registers.xlsx \
    --manufacturer "TestMfg" \
    --model "TEST-2000" \
    -o test_excel_output.csv
```

## Troubleshooting

### Problem: No registers extracted
- Run with `-v` (verbose) to see detection logs.
- Ensure the source file has clear column headers.
- Use `--mapping` if the column names are non-standard.

### Problem: Invalid addresses
- Check if you need an `--address-offset`.
- Verify the address column is correctly detected in verbose mode.

## What You Get

✅ Properly formatted WebdynSunPM header
✅ Indexed register entries with O(1) memory efficiency
✅ Automatic data type mapping and normalization
✅ Address overlap and duplicate detection
✅ Scaling factor calculation

**Ready to use with WebdynSunPM!**
