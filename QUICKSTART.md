# WebdynSunPM DefFileGenerator — Quick Start Guide

<p align="center">
  <a href="README.md"><b>Home & Installation</b></a> •
  <a href="docs/input-format.md"><b>Input Formats</b></a> •
  <a href="docs/architecture.md"><b>Architecture</b></a> •
  <a href="DEVELOPER_GUIDE.md"><b>Developer Guide</b></a> •
  <a href="docs/security.md"><b>Security</b></a>
</p>

---

> [!NOTE]
> For a comprehensive, step-by-step installation guide designed for complete beginners, check out the [A-to-Z Installation Guide in the README](README.md#-a-to-z-installation-guide-absolute-beginner-friendly).

---

## 📑 Table of Contents

- [🎯 What This Tool Does](#-what-this-tool-does)
- [⚡ Quick Installation Options](#-quick-installation-options)
- [🚀 Real-World Conversion Examples](#-real-world-conversion-examples)
  - [Example 1: Inverter PDF Datasheet (GoodWe)](#example-1-inverter-pdf-datasheet-goodwe)
  - [Example 2: Multi-Sheet Excel Workbook (SMA)](#example-2-multi-sheet-excel-workbook-sma)
  - [Example 3: Exported CSV Register Map (ABB)](#example-3-exported-csv-register-map-abb)
- [🔍 How Conversion Works Under the Hood](#-how-conversion-works-under-the-hood)
- [🛠️ Command-Line Options Reference](#️-command-line-options-reference)
- [🧪 Testing Immediately with Bundled Sample Files](#-testing-immediately-with-bundled-sample-files)
- [✅ What the Generated WebdynSunPM File Contains](#-what-the-generated-webdynsunpm-file-contains)
- [💡 Pro-Tips for Best Results](#-pro-tips-for-best-results)
- [❓ Quick Troubleshooting](#-quick-troubleshooting)

---

## 🎯 What This Tool Does

This tool **automatically extracts** Modbus register tables from heterogeneous manufacturer documentation files and generates clean, verified definition files for **WebdynSunPM** telemetry gateways.

Simply supply a source document (**PDF, Excel `.xlsx`/`.xlsm`, CSV, or XML**), and the engine will:
1. Locate register tables automatically across pages or sheets.
2. Extract addresses, names, data types, units, and scaling factors.
3. Validate Modbus constraints (boundaries, overlaps, bitfield slices).
4. Produce a standardized semicolon-delimited CSV definition file ready to be uploaded to your gateway.

---

## ⚡ Quick Installation Options

### Option A: Pre-compiled Windows Standalone Executable (.exe)
1. Download `DefFileGenerator-vX.X.X-windows-x64.zip` from the [GitHub Releases](https://github.com/Cylae/DefFileGenerator/releases).
2. Right-click the ZIP archive > **Extract All...**.
3. Double-click **`DefFileGenerator-GUI.exe`** to launch the Windows 11 desktop app.

### Option B: Python Installation via Pip
```bash
# Core CLI and Windows 11 Desktop GUI:
pip install -e .

# Or with FastAPI Web Interface:
pip install -e ".[web]"
```

---

## 🚀 Real-World Conversion Examples

### Example 1: Inverter PDF Datasheet (GoodWe)

Suppose you have a manufacturer PDF datasheet containing a table like this:

```text
Modbus Register Map
-------------------
Register | Parameter Name    | Type   | Unit | Scale | Access
---------|-------------------|--------|------|-------|-------
40001    | AC Power          | uint16 | W    | 1     | R
40002    | DC Voltage        | uint16 | V    | 0.1   | R
40003    | Temperature       | int16  | °C   | 0.1   | R
```

Run:
```bash
deffilegen run "GoodWe_Modbus_Datasheet.pdf" \
    --manufacturer "GoodWe" \
    --model "GW5000-DNS" \
    -o "goodwe_definition.csv"
```

The output file `goodwe_definition.csv` is generated instantly:
```csv
modbusRTU;Inverter;GoodWe;GW5000-DNS;;;;;;;
1;3;40001;U16;;AC Power;ac_power;1.000000;0.000000;W;4
2;3;40002;U16;;DC Voltage;dc_voltage;0.100000;0.000000;V;4
3;3;40003;I16;;Temperature;temperature;0.100000;0.000000;°C;4
```

---

### Example 2: Multi-Sheet Excel Workbook (SMA)

To convert an Excel workbook containing registers:

```bash
# Process all worksheets in the workbook
deffilegen run "SMA_Inverter_Registers.xlsx" \
    --manufacturer "SMA" \
    --model "STP-5000TL" \
    -o "sma_definition.csv"

# Or target a specific worksheet
deffilegen run "SMA_Inverter_Registers.xlsx" \
    --sheet "Holding Registers" \
    --manufacturer "SMA" \
    --model "STP-5000TL" \
    -o "sma_definition.csv"
```

---

### Example 3: Exported CSV Register Map (ABB)

To convert a CSV export from a manufacturer configuration tool:

```bash
deffilegen run "abb_registers.csv" \
    --manufacturer "ABB" \
    --model "PVS-5.0-TL" \
    -o "abb_definition.csv"
```

---

## 🔍 How Conversion Works Under the Hood

```text
[ Source File ]      ──▶  [ Step 1: Extraction ]  ──▶  [ Step 2: Normalization ]  ──▶  [ Step 3: Webdyn Output ]
(PDF, XLSX, CSV, XML)     Smart column matching        Decimal / Hex addresses          UTF-8 with BOM CSV
                          (address, name, type...)     Canonical types (U16, F32...)    Semicolon-delimited
                                                       Linear scaling & unique tags     Overlap-free & validated
```

### Step 1: Intelligent Column Recognition
The extractor recognizes headers in English, French, and German:
- **Address**: `address`, `addr`, `register`, `reg`, `adresse`, `offset`, `index`
- **Name**: `name`, `description`, `parameter`, `variable`, `signal`, `grandeur`
- **Type**: `type`, `data type`, `datatype`, `format`, `taille`
- **Unit**: `unit`, `units`, `unite`, `symbole`
- **Scale / Factor**: `scale`, `factor`, `multiplier`, `gain`, `ratio`, `pas`

### Step 2: Data Normalization
- Hexadecimal addresses (`0x9C40` or `9C40h`) are automatically converted to decimal (`40000`).
- Vendor type expressions (`uint16`, `float32`, `signed int`, etc.) are mapped to canonical Webdyn codes (`U16`, `F32`, `I16`, etc.).
- Bitfield slices (`address_startbit_length`) and character strings (`address_length`) are normalized.
- Clean, lowercase variable tags are synthesized automatically.

### Step 3: Strict Validation & Atomic Disk Writing
The definition is validated by the $O(\log N)$ interval engine (duplicate tag checks, bit slice limits, memory collisions) before being atomically written to disk with `os.fsync` and `os.replace`.

---

## 🛠️ Command-Line Options Reference

General syntax:
```bash
deffilegen run SOURCE_FILE --manufacturer BRAND --model MODEL [OPTIONS]
```

### Mandatory Arguments
| Option | Description | Example |
|---|---|---|
| `SOURCE_FILE` | Path to manufacturer documentation (PDF, Excel, CSV, or XML) | `datasheet.pdf` |
| `--manufacturer` | Equipment brand or manufacturer name | `"Huawei"` |
| `--model` | Equipment model name | `"SUN2000-50KTL"` |

### Optional Arguments
| Option | Default | Description |
|---|---|---|
| `-o, --output` | Auto-generated | Path of the generated CSV file |
| `--protocol` | `modbusRTU` | Modbus protocol (`modbusRTU` or `modbusTCP`) |
| `--category` | `Inverter` | Equipment classification (`Inverter`, `Meter`, `Sensor`, `Battery`) |
| `--sheet` | *(all)* | Target specific Excel worksheet |
| `--pages` | *(all)* | Target specific PDF pages (e.g. `1,3-5` or `12-18`) |
| `--mapping` | *(heuristic)* | Path to custom JSON column mapping override file |
| `--address-offset` | `0` | Global address shift (e.g. `-1` to fix 1-based indexing) |
| `--forced-write` | `""` | Forced writing code inserted into CSV header |
| `--force` | `false` | Overwrite destination file without prompting |
| `--no-validate` | `false` | Skip the default post-generation validation step |
| `-v, --verbose` | `false` | Enable verbose diagnostic logging |

---

## 🧪 Testing Immediately with Bundled Sample Files

The repository bundles sample files so you can verify your installation immediately:

### Test 1: Using the Bundled Excel File
```bash
deffilegen run sample_inverter_registers.xlsx \
    --manufacturer "TestBrand" \
    --model "TEST-2000" \
    -o test_excel_output.csv
```

### Test 2: Using the Bundled CSV Template
```bash
deffilegen run template.csv \
    --manufacturer "TestBrand" \
    --model "TEST-1000" \
    -o test_csv_output.csv
```

Audit the resulting definition file:
```bash
deffilegen validate test_excel_output.csv
```

---

## ✅ What the Generated WebdynSunPM File Contains

The resulting CSV file strictly satisfies WebdynSunPM telemetry gateway specifications:

1. **Header Row (Line 1):**
   ```text
   modbusRTU;Inverter;GoodWe;GW5000-DNS;;;;;;;
   ```
2. **Register Rows (Lines 2+) with exactly 11 semicolon-separated columns:**
   - **Index**: 1-based sequential counter (`1`, `2`, `3`...).
   - **Info1**: Modbus function code (`1` = Coil, `2` = Discrete Input, `3` = Holding Register, `4` = Input Register).
   - **Info2**: Normalized Modbus address (e.g. `40001`, `30001_0_1` for bits, `30030_20` for strings).
   - **Info3**: Normalized Webdyn data type (`U16`, `I16`, `U32`, `F32`, `STRING`, `BITS`, etc.).
   - **Info4**: Reserved field (always empty `""`).
   - **Name**: Clear human-readable description.
   - **Tag**: Unique variable identifier in lowercase ASCII.
   - **CoefA**: Multiplier factor formatted to 6 decimal places ($Factor \times 10^{ScaleFactor}$).
   - **CoefB**: Offset / bias formatted to 6 decimal places.
   - **Unit**: Engineering physical unit (`V`, `A`, `W`, `°C`, `Hz`).
   - **Action**: Webdyn action code (`4` for normal read, `10` for constant).

---

## 💡 Pro-Tips for Best Results

1. **Prefer text-based documents**: If you can select and highlight text in the PDF with your mouse, the extractor will read tables with 100% fidelity.
2. **Narrow down PDF pages**: For large 300-page user manuals where registers only appear on pages 45 to 52, use `--pages 45-52` to save time.
3. **Use verbose mode for diagnostics**: If column detection seems unexpected, pass `-v` to inspect exact heuristic scores.
4. **Always validate before site deployment**: Run `deffilegen validate output.csv` to ensure zero register collisions before importing into your gateway.

---

## ❓ Quick Troubleshooting

- **No registers found in PDF?**  
  The PDF might be a scanned image. Use an OCR utility to convert scanned pages into selectable text first.
- **Addresses are off by 1?**  
  Some vendors use 0-based addresses while others use 1-based addresses. Use `--address-offset 1` or `--address-offset -1` to align them.
- **Unusual vendor column headers?**  
  Provide a simple `mapping.json` file (e.g. `{"Address": "Reg_Num", "Name": "Signal"}`) and pass `--mapping mapping.json`.

For the complete troubleshooting matrix, check the [Troubleshooting FAQ in the README](README.md#-beginner-sos--troubleshooting-faq).

---

<p align="center">
  <a href="README.md"><b>⬅ Back to README</b></a> •
  <a href="docs/input-format.md"><b>Input Formats Guide ➡</b></a>
</p>
