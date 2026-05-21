# WebdynSunPM Documentation Parser - Quick Start Guide

## What This Tool Does

This tool **automatically extracts** register information from manufacturer documentation files and generates WebdynSunPM definition files.

Simply provide a PDF, Excel, CSV, or XML file from the manufacturer, and it will:
1. Find the register tables
2. Extract addresses, names, data types, units, etc.
3. Generate a ready-to-use WebdynSunPM definition file

## Installation

### Core Dependencies (Required for basic execution)
```bash
pip install openpyxl pdfplumber defusedxml lxml reportlab
```

### Optional Dependencies (Required for stress testing and performance)
```bash
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

### From CSV File
```bash
python doc_to_webdyn.py registers.csv \
    --manufacturer "Fronius" \
    --model "Symo-5.0" \
    -o fronius_definition.csv
```

### From XML File
```bash
python doc_to_webdyn.py registers.xml \
    --manufacturer "Generic" \
    --model "Model-X" \
    -o generic_definition.csv
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

- **Normalization**: Handles hex addresses (e.g., `0x9C40`) and decimal automatically.
- **Type Mapping**: Converts `uint16` -> `U16`, `float32` -> `F32`, etc.
- **Tag Generation**: Creates unique alphanumeric tags from register names.
- **Coefficients**: Calculates `CoefA` and `CoefB` using factors and scale factors.

### Step 3: Creates WebdynSunPM definition file

Outputs a properly formatted CSV file ready for WebdynSunPM:
```csv
modbusRTU;Inverter;Huawei;SUN2000-5KTL;;;;;;;
1;3;40001;U16;;Active Power;active_power;1.000000;0.000000;W;4
2;3;40002;U16;;Voltage;voltage;0.100000;0.000000;V;4
...
```

## Examples with Real Files

### Example 1: PDF Datasheet with Register Tables

If you have a PDF like this:

```
Modbus Register Map
-------------------
Register | Parameter Name    | Type   | Unit | Scale | Access
---------|-------------------|--------|------|-------|-------
40001    | AC Power         | uint16 | W    | 1     | R
40002    | DC Voltage       | uint16 | V    | 0.1   | R
40003    | Temperature      | int16  | °C   | 0.1   | R
```

Run:
```bash
python doc_to_webdyn.py inverter_datasheet.pdf \
    --manufacturer "GoodWe" \
    --model "GW5000-DNS" \
    -o goodwe_definition.csv
```

Output:
```csv
modbusRTU;Inverter;GoodWe;GW5000-DNS;;;;;;;
1;3;40001;U16;;AC Power;ac_power;1.000000;0.000000;W;4
2;3;40002;U16;;DC Voltage;dc_voltage;0.100000;0.000000;V;4
3;3;40003;I16;;Temperature;temperature;0.100000;0.000000;°C;4
```

### Example 2: Excel File

If you have an Excel file with sheets containing register information:

```bash
# Process all sheets
python doc_to_webdyn.py Inverter_Registers.xlsx \
    --manufacturer "SMA" \
    --model "STP-5000TL" \
    -o sma_definition.csv

# Process specific sheet
python doc_to_webdyn.py Inverter_Registers.xlsx \
    --sheet "Holding Registers" \
    --manufacturer "SMA" \
    --model "STP-5000TL" \
    -o sma_definition.csv
```

### Example 3: CSV Export from Manufacturer Tool

If you exported a CSV from the manufacturer's software:

```bash
python doc_to_webdyn.py exported_registers.csv \
    --manufacturer "ABB" \
    --model "PVS-5.0-TL" \
    -o abb_definition.csv
```

## Command-Line Options

```bash
python doc_to_webdyn.py INPUT_FILE --manufacturer MFG --model MODEL [OPTIONS]
```

### Required Arguments
- `INPUT_FILE` - Your PDF, Excel, CSV, or XML file
- `--manufacturer MFG` - Manufacturer name (e.g., "Huawei")
- `--model MODEL` - Model name (e.g., "SUN2000-5KTL")

### Optional Arguments
- `-o OUTPUT` - Output filename (default: auto-generated)
- `--protocol PROTO` - Protocol name (default: modbusRTU)
- `--category CAT` - Device category (default: Inverter)
- `--sheet NAME` - Excel sheet name (processes all if not specified)
- `--pages LIST` - Comma-separated list of PDF pages to process
- `--mapping FILE` - JSON mapping file for custom column detection
- `--address-offset VAL` - Integer value to add to all addresses
- `--forced-write VAL` - Value for the forced-write column in the header
- `-v, --verbose` - Show detailed processing information

## Troubleshooting

### Problem: No registers extracted

**Solution:**
1. Check if your file has clearly labeled columns
2. Run with `-v` (verbose) to see what's happening:
   ```bash
   python doc_to_webdyn.py yourfile.pdf --manufacturer "X" --model "Y" -v
   ```
3. Make sure tables in PDF are text-based (not scanned images)

### Problem: Wrong data types

**Solution:**
- Add a "Type" or "Data Type" column to your source file
- The tool will guess if not specified

### Problem: Incorrect addresses

**Solution:**
- Check if addresses are in the right column
- The tool handles hex (0x9C40) and decimal (40001) automatically
- Use `--address-offset` if addresses are shifted.

## Tips for Best Results

1. **Start with clean documentation** - Well-formatted source files work best
2. **Test first** - Try with sample files to understand the output
3. **Use verbose mode** - Add `-v` to see what's being detected
4. **Review output** - Always check the generated file
5. **Keep originals** - Save your source documentation for reference

---

**You're ready to go! Just point the tool at your manufacturer documentation and it will do the rest.**
