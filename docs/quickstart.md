# Quick Start Cheatsheet & CLI Reference

<p align="center">
  <a href="../README.md"><b>README</b></a> •
  <a href="../QUICKSTART.md"><b>Beginner Tutorial</b></a> •
  <a href="input-format.md"><b>Input Formats</b></a> •
  <a href="architecture.md"><b>Architecture</b></a> •
  <a href="security.md"><b>Security</b></a>
</p>

---

## ⚡ Direct File Conversion

Convert a manufacturer documentation file into a WebdynSunPM definition CSV in a single step:

```bash
deffilegen run register_map.xlsx \
  --manufacturer "Huawei" \
  --model "SUN2000-50KTL" \
  -o huawei_definition.csv
```

Supported input formats: **PDF, Excel (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`), CSV, and XML**.

---

## 🛠️ Multi-Step Pipeline Commands

For custom workflows and intermediate data inspection, `deffilegen` provides decoupled subcommands:

```text
┌─────────────────┐       ┌─────────────────┐       ┌─────────────────┐
│   1. extract    │ ───▶  │   2. generate   │ ───▶  │   3. validate   │
│ (Raw registers) │       │  (Webdyn CSV)   │       │(Strict checks)  │
└─────────────────┘       └─────────────────┘       └─────────────────┘
```

```bash
# 1. Extract raw registers into an intermediate CSV file
deffilegen extract datasheet.pdf -o raw_registers.csv

# 2. Generate a WebdynSunPM definition CSV from intermediate registers
deffilegen generate raw_registers.csv --manufacturer "SMA" --model "STP5000" -o sma_def.csv

# 3. Validate an existing definition file
deffilegen validate sma_def.csv

# Optional: Validate in lenient mode (permits address overlaps for diagnostics)
deffilegen validate sma_def.csv --lenient
```

---

## 📋 Common Command Options

| Flag | Description | Example |
|---|---|---|
| `-o, --output FILE` | Destination output file path | `-o definition.csv` |
| `--protocol PROTO` | Modbus protocol (`modbusRTU` or `modbusTCP`) | `--protocol modbusTCP` |
| `--category CAT` | Device classification (`Inverter`, `Meter`, `Sensor`, `Battery`) | `--category Inverter` |
| `--sheet NAME` | Specific Excel worksheet name | `--sheet "Holding Registers"` |
| `--pages RANGE` | Selective PDF pages range | `--pages "12,14-18"` |
| `--mapping FILE` | Custom JSON column mapping dictionary | `--mapping custom_map.json` |
| `--address-offset N` | Global address shift (applied once) | `--address-offset -1` |
| `--force` | Overwrite existing output without prompt | `--force` |
| `--no-validate` | Skip post-generation validation step | `--no-validate` |
| `-v, --verbose` | Enable diagnostic logging output | `-v` |

---

## 📄 Generating Built-in Templates

Create blank template files without requiring an input document:

```bash
# Generate intermediate register input template
deffilegen generate --template --template-mode input -o template_input.csv

# Generate full WebdynSunPM definition CSV template
deffilegen generate --template --template-mode definition -o template_definition.csv
```

---

## 🧪 Quality Gates & Validation Suite

```bash
# Run complete test suite (685 tests)
pytest

# Code formatting and linting
ruff check .
ruff format --check .

# Static type analysis
mypy DefFileGenerator web

# Security scanning
bandit -r DefFileGenerator web -ll
```

---

<p align="center">
  <a href="../README.md"><b>⬅ Back to README</b></a> •
  <a href="input-format.md"><b>Input Formats Guide ➡</b></a>
</p>
