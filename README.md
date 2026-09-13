# WebdynSunPM DefFileGenerator & Documentation Parser

A production-grade Python library, CLI, and optional FastAPI interface for extracting Modbus register maps from heterogeneous manufacturer documentation (**PDF, Excel, CSV, XML**) and generating validated **WebdynSunPM definition CSV files**.

## Architecture

```text
[ PDF / Excel / CSV / XML ]
            │
            ▼
[ Extractor ]
  format-specific readers
  + lazy row generators
            │
            ▼
[ Mapper / Cleaner ]
  exact → synonym → partial
  column resolution
            │
            ▼
[ Generator Core ]
  types · addresses · tags
  coefficients · actions
            │
            ▼
[ Validation Layer ]
  range · tag uniqueness
  register + bit overlap
            │
            ▼
[ Atomic CSV Writer ]
  sanitization · fsync
  atomic replace
```

### Extraction

- **PDF** — `pdfplumber` table extraction with page selection and newline cleanup.
- **Excel** — `openpyxl` in `read_only=True`, `data_only=True` mode.
- **CSV** — BOM-aware decoding plus delimiter detection.
- **XML** — `defusedxml.ElementTree`, rejecting DTD/entity based attacks.

### Column mapping

The mapper resolves vendor-specific columns in three tiers:

1. case-insensitive exact internal-name match;
2. exact synonym match against `Extractor.COLUMN_MAPPING`;
3. partial fallback match.

Explicit JSON overrides remain available through `--mapping`.

### Modbus semantics

The generator normalizes:

- decimal and hexadecimal addresses;
- register widths for `U8/I8`, `U16/I16`, `U32/I32/F32/IP`, `U64/I64/F64`, `MAC`, `IPV6`, `STRING`, and `BITS`;
- data type / endianness aliases;
- coefficients (`CoefA = Factor × 10^ScaleFactor`, `CoefB = Offset`);
- tags, register type codes, and actions.

`BITS` values use `address_startbit_length`. A slice must stay inside one 16-bit register, and strict validation detects overlapping slices on the same base register. See [`docs/input-format.md`](docs/input-format.md).

## Requirements & installation

Python **3.10+** is required.

```bash
# Core CLI / library
python -m pip install -e .

# Core + FastAPI web interface
python -m pip install -e ".[web]"

# Development / test dependencies
python -m pip install -e ".[dev,web]"
```

The wheel includes the `web` package and its static frontend assets; CI builds and inspects the wheel to enforce that packaging contract.

## CLI

The installed entry point is `deffilegen`.

```bash
# Full extraction + generation
deffilegen run datasheet.pdf \
  --manufacturer "Huawei" \
  --model "SUN2000" \
  -o huawei_sun2000.csv

# Extraction only
deffilegen extract inverter_map.xlsx \
  --sheet "Registers" \
  -o extracted.csv

# Generate from normalized CSV
deffilegen generate extracted.csv \
  --manufacturer "SMA" \
  --model "STP5000" \
  -o sma_stp5000.csv

# Strict validation
deffilegen validate sma_stp5000.csv
```

Main subcommands:

- `run` — extract, generate, and validate in one workflow;
- `extract` — create an intermediate normalized register CSV;
- `generate` — build a WebdynSunPM definition from normalized data;
- `validate` — check an existing definition.

## Programmatic API

```python
from generate_webdyn_def import generate_webdyn_definition

success = generate_webdyn_definition(
    input_file="registers.xlsx",
    output_file="output_definition.csv",
    manufacturer="SMA",
    model="STP5000",
    protocol="modbusRTU",
    category="Inverter",
    address_offset=0,
    strict_validation=True,
)
```

## Web interface

Run the optional FastAPI backend with:

```bash
uvicorn web.app:app --host 127.0.0.1 --port 8000
```

Endpoints:

- `GET /api/health`
- `POST /api/convert`
- `POST /api/validate`

Uploads are streamed in bounded chunks and capped at **10 MiB**. Parser/generator exceptions are logged server-side while API clients receive sanitized error messages. Wildcard CORS is non-credentialed.

## Security & integrity

Key controls include:

- **CSV formula-injection mitigation** for ASCII and Unicode trigger variants;
- **control-character filtering** before CSV output;
- **XXE / entity-expansion protection** through `defusedxml`;
- **atomic output replacement** so a failed generation cannot truncate an existing definition file;
- **bit-slice range and overlap validation** for `BITS` addresses;
- **bounded web uploads** and non-disclosure of internal exception text;
- **privileged auto-merge trust checks** restricting the Jules workflow to owner-authored, same-repository PRs after successful CI.

See [`docs/security.md`](docs/security.md) for the complete security model.

## Quality gates

The established suite contains **500+ unit and integration tests**. GitHub Actions runs the project across Python 3.10, 3.11, and 3.12 with:

```text
ruff check .
ruff format --check .
mypy DefFileGenerator web
bandit -r DefFileGenerator web -ll
python -m pytest --cov=DefFileGenerator --cov=web
```

CI also builds a wheel and verifies that `web/app.py` plus the static frontend assets are present in the distributable artifact.

Additional stress batteries cover ambiguous column resolution, large register maps, string length semantics, and type normalization behavior.

## Documentation

- [`docs/quickstart.md`](docs/quickstart.md) — concise usage walkthrough
- [`docs/input-format.md`](docs/input-format.md) — accepted columns and address notation
- [`docs/architecture.md`](docs/architecture.md) — internal architecture
- [`docs/security.md`](docs/security.md) — security and integrity guarantees
- [`docs/development.md`](docs/development.md) — contributor workflow
- [`docs/booklet/def-file-generator-booklet.html`](docs/booklet/def-file-generator-booklet.html) — versioned source for the visual Canva booklet

## License

MIT.
