# Quick Start Guide

## Installation

Python 3.10 or newer is required. To install the CLI and desktop application in editable mode:

```bash
pip install -e .
```

To install development and quality assurance dependencies:

```bash
pip install -e ".[dev,web]"
```

## Direct File Conversion

Transform a manufacturer register documentation file into a WebdynSunPM definition CSV:

```bash
deffilegen run register_map.xlsx \
  --manufacturer "Webdyn" \
  --model "DeviceModel" \
  -o definition.csv
```

Supported inputs are PDF, CSV, XML, and modern Excel files (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`). Legacy `.xls` files must first be converted to `.xlsx`.

## Multi-Step Workflow CLI

Use the main command-line interface (`deffilegen`) to separate register extraction, generation, and validation steps:

```bash
# 1. Extract registers to an intermediate CSV
deffilegen extract register_map.xlsx -o registers.csv

# 2. Generate the WebdynSunPM definition file
deffilegen generate registers.csv --manufacturer "Webdyn" --model "DeviceModel" -o definition.csv

# 3. Validate the definition file
deffilegen validate definition.csv
```

Useful extraction controls:

```bash
# Select PDF pages and overwrite an existing output intentionally
deffilegen run manual.pdf --pages "12,14-18" --force \
  --manufacturer "Vendor" --model "Model" -o definition.csv

# Select an Excel worksheet and provide explicit source-column names
deffilegen extract map.xlsx --sheet "Holding Registers" \
  --mapping mapping.json -o registers.csv
```

Use `--address-offset` only when the manufacturer's addressing convention is known. The offset is applied once during extraction. `run` validates its result by default; `--no-validate` disables that final check. `validate --lenient` permits overlaps but still checks the file structure and field syntax.

Image-only PDFs require OCR before extraction. A successful command means the recovered rows are valid; it does not prove that an unreadable source table was complete. Compare register counts and boundaries with the manufacturer's manual.

Run `deffilegen --help` or `deffilegen <subcommand> --help` to inspect all available arguments.

## Quality Checks Before Delivery

```bash
uv run --all-extras pytest --cov=DefFileGenerator --cov=web
uv run ruff check .
uv run ruff format --check .
uv run mypy DefFileGenerator web
uv run bandit -r DefFileGenerator web -ll
uv build
```
