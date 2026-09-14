# Changelog

All notable changes to this project are documented in this file.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and the project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.1]

Rebuilt against the current `main`. Functionally audited, hardened, and verified across both Core engine and Web service components.

### Security

- **Web API Path Traversal Defense.** Hardened `/api/convert` and `/api/validate` endpoints in `web/app.py` using `os.path.basename` to prevent directory traversal sequence attacks (e.g., `../../etc/passwd`).
- **Core Error Isolation in Web Service.** Wrapped core extraction, generator, and validation calls in explicit `try...except` blocks to safely trap core exceptions and return structured HTTP 400 JSON responses without exposing internal stack traces.
- **CSV Line-Break Injection Defense.** Enhanced `sanitize_csv_field` in `def_gen.py` to convert embedded carriage returns (`\r\n`, `\r`, `\n`) into single spaces, preventing multiline table cell descriptions from splitting and corrupting WebdynSunPM definition row syntax.
- **Web Frontend XSS Hardening.** Escaped single quotes (`'`) to `&#39;` in `web/static/app.js` `escapeHtml` to prevent script execution vectors via dynamically injected attributes.
- **Formula-injection sanitiser hardened.** `sanitize_csv_field` previously
  inspected only the raw first byte. Spreadsheet clients strip leading
  whitespace before deciding whether a cell is a formula and normalise
  full-width punctuation to ASCII, so payloads prefixed with a tab, CR, LF,
  vertical tab, form feed or NBSP — and full-width `＝ ＋ － ＠` — were written
  unescaped. The DDE pipe (`|`) was not treated as a trigger at all. Escaping
  is now decided on the first *significant* character. The numeric bypass
  additionally requires `math.isfinite()` and a strict decimal pattern, so
  `-inf`, `+nan` and `-1_000` are escaped while `-10.5`, `+25` and `1.5e3`
  remain verbatim. This closes a path from untrusted vendor documentation to
  code execution on an engineer's workstation.

### Fixed

- **Windows Openpyxl File Descriptor Locks.** Eliminated Windows `PermissionError: [WinError 32]` during temporary file cleanup by switching `Extractor.extract_from_excel` from `load_workbook(..., read_only=True)` to in-memory `io.BytesIO` parsing.
- **Multi-Encoding & Real-World CSV Ingestion.** Added multi-stage fallback decoding (`utf-16` BOM, `utf-8`, `cp1252`, `errors="replace"`) and automatic detection of WebdynSunPM metadata headers (`modbus;...`) in `Extractor.extract_from_csv`.
- **Web Filename Fallbacks.** Added fallback defaults (`"manufacturer"`, `"model"`) when sanitized parameters are empty or whitespace-only.

- **Non-reproducible output.** Column detection iterated a `set` of header
  names. Because string hashing is randomised per interpreter, targets sharing
  a pattern (`offset` matches both Address and Offset) could bind to different
  source columns between runs — the same input produced different definition
  files, and rows were silently dropped on some seeds. Header order is now
  preserved.
- **Action code `10` was discarded.** `Constant` variables (firmware 5.2.02+)
  were rewritten to `1`, changing device behaviour. The code is now preserved,
  with the fallback for genuinely unknown codes unchanged.
- **Non-atomic writes.** A failure mid-stream truncated an existing definition
  file. Output is now staged via `tempfile.mkstemp` on the same filesystem and
  moved into place with `os.replace`; no temporary file survives a failure.

### Added

- Post-generation validation for `run`, with `--no-validate` to opt out.
- `--force` guard: an existing output file is no longer silently overwritten.
- `--lenient` for `validate`, downgrading address overlaps to warnings.
- `-q/--quiet` and `--version`; `-v`/`-q` are accepted before *or* after the
  sub-command.
- Actionable diagnostics: unsupported extensions list the supported formats,
  missing flags are named with a worked example, and empty extractions suggest
  `--sheet`, `--pages` or `--mapping`.
- Expanded web unit test coverage in `DefFileGenerator/tests/test_web.py` to 11 tests covering path traversal sanitization, corrupt register uploads, definition validation checks, and temporary file lifecycle on Windows.
- Real-world manufacturer validation battery (`DefFileGenerator/tests/run_equipementiers_battery.py`) validating extraction and definition generation across 9,862 real registers from equipment manufacturer files (.xlsx, .csv, .xml).
- Dedicated 30-manufacturer benchmark suite (`DefFileGenerator/tests/test_30_equipementiers.py`) achieving 30/30 (100%) success across 10 industry manufacturers: Huawei, ABB, Fox ESS, GoodWe, Growatt, Hukseflux, Janitza, Kaco, Lettel, and Siebert.
- **Native Windows 11 Desktop Application (`DefFileGenerator/gui.py`).**
  - High-DPI per-monitor awareness via Windows 11 API `SetProcessDpiAwareness(2)`.
  - Responsive multi-threaded worker architecture (`threading.Thread`) keeping the GUI reactive without freezing.
  - Dark/Light mode theme integration matching Windows 11 system preferences.
  - 4 functional tabs: *Conversion & Génération* (with live register table preview), *Validateur de Définition* (line-by-line syntax & overlap inspection), *Modèles d'Équipement* (pre-configured Inverters, Meters, Pyranometers, Batteries, Weather Stations), and *Console & Logs*.
  - Entry points: `deffilegen-gui`, `deffilegen gui`, `deffilegen --gui`, and double-click `launch_gui.bat`.
- **Smart Header Detection for PDF Tables.** Added semantic keyword density scanning across the first 4 rows of PDF tables in `Extractor.extract_from_pdf` to bypass merged title rows.
- **Structured Validation Diagnostics.** Added `validate_csv_detailed` returning `ValidationReport` and `ValidationIssue` objects for rich diagnostics across CLI, Web API, and Desktop GUI.
- Smart Excel Header Detection in `Extractor.extract_from_excel`: dynamically scans the first 15 rows of each worksheet to identify actual Modbus table headers, eliminating silent column misalignment caused by title, metadata, or section banners.
- International Column Synonyms in `Extractor.COLUMN_MAPPING`: added SunSpec and multilingual variants (`point`, `point name`, `designation`, `nom`, `grandeur`, `quantity`, `adresse`, `start reg (dec)`, `registre`, `res.`, `resolution`, `accès`, `acces`, `lecture/écriture`).
- Enhanced Webdyn definition CSV ingestion in `Extractor.extract_from_csv`: supports standard 4-part headers (`protocol;category;manufacturer;model`) as well as 5+ parts.
- Enhanced strict validation in `validate_csv` verifying `Info1` in `(1, 2, 3, 4)`, valid WebdynSunPM Action codes, finite float precision for `CoefA`/`CoefB`, and minimum column length.
- Added `info1`, `info2`, `info3`, `coefa`, and `coefb` aliases to `COLUMN_MAPPING` in `DefFileGenerator/extractor.py`.

### Changed

- Register processing throughput improved ~1.46x on a 50,000-register map.
  `normalize_type` is substantially faster: twelve sequential `re.search` calls
  per row were replaced by a precompiled, specificity-ordered table behind an
  `lru_cache`. Address normalisation and type-width lookup are likewise
  memoised, action codes moved to `frozenset`, and `bisect` was hoisted out of
  the hot loop.
- Column detection reduced from O(K x T x P) to O(K) by lowercasing each header
  once instead of per probe.
