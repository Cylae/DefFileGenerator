# Quick Start Guide

## Installation

To install the package in editable mode:

```bash
pip install -e .
```

To install development and quality assurance dependencies:

```bash
pip install -e ".[dev]"
```

## Direct File Conversion

Transform a manufacturer register documentation file into a WebdynSunPM definition CSV:

```bash
python doc_to_webdyn.py register_map.xlsx \
  --manufacturer "Webdyn" \
  --model "DeviceModel" \
  -o definition.csv
```

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

Run `deffilegen --help` or `deffilegen <subcommand> --help` to inspect all available arguments.

## Quality Checks Before Delivery

```bash
pytest
ruff check .
mypy DefFileGenerator generate_webdyn_def.py doc_to_webdyn.py
```
