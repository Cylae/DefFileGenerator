# Changelog

All notable changes to this project are documented in this file.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and the project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Performance

- **CSV Field Sanitization Fast-Path (3.09x speedup).** Added an early ASCII alphanumeric short-circuit check (`s.isascii() and s.isprintable() and s[0].isalnum()`) in `Generator.sanitize_csv_field` avoiding character-by-character generator comprehensions, surrogate scans, and string replacements for >95% of standard Modbus table cells while preserving 100% defense against CSV formula injection (DDE attacks).
- **Cached & Short-Circuited Address Normalization (4.48x speedup).** Implemented `@functools.lru_cache(maxsize=4096)` backed helper `_normalize_address_cached` and an immediate decimal integer fast-path in `Generator.normalize_address_val`, making repetitive address parsing during mapping and validation instant $O(1)$.
- **Direct Dictionary & Set Type Lookups (1.50x speedup).** Added `_EXACT_REG_COUNTS`, `_COMMON_VALID_NUMERIC`, `_VALID_SPECIAL_TYPES`, `_CANONICAL_TYPE_SET`, and `_SHORTHAND_TYPE_MAP` to replace sequential linear tuple scans and regular expressions with $O(1)$ hash table lookups during type validation and register count calculation.
- **End-to-End Definition Generation Speedup.** Reduced processing time for 5,000-register CSV definition generation from 0.356s to 0.235s (**34% faster**) and address offset stress processing from 0.348s to 0.212s (**39% faster**) in `run_gigantic_battery.py`.

### Added

- **Wave 10 Hardening Regression Suite (`DefFileGenerator/tests/test_wave10_hardening.py`).** Added permanent automated tests covering surplus CSV column handling, ragged rows, GUI CLI flags, and builder environment discovery.
- **Standalone Windows Executables (.exe).** Added reproducible PyInstaller specification (`DefFileGenerator.spec`), build orchestration script (`build_exe.py`), and one-click Windows launcher (`build_exe.bat`) compiling `DefFileGenerator-GUI.exe` (windowed desktop UI) and `deffilegen.exe` (portable CLI).
- **Automated GitHub Actions Build & Release Workflow (`.github/workflows/build-exe.yml`).** Automatically compiles and smoke-tests Windows binaries on `windows-latest` runners on push to `main` (published as downloadable workflow artifacts) and on release tags `v*` (published as GitHub Release assets with zip bundles and SHA-256 checksums).
- **Build System Unit Tests (`DefFileGenerator/tests/test_build_exe.py`).** Added automated test coverage validating spec targets, package collections, CLI argument parsing, SHA-256 checksum generation, and workflow syntax.
- **Developer Onboarding Manual (`DEVELOPER_GUIDE.md`).** Detailed technical documentation describing system architecture, data pipeline, WebdynSunPM domain rules, security invariants, developer extension recipes, and validation tooling.
- **Clarity & Efficiency Benchmark Suite (`DefFileGenerator/tests/test_clarity_and_efficiency.py`).** 9 automated tests verifying domain constants, pre-compiled regex tables, frozen keyword sets, documentation coverage, and sub-millisecond type normalization performance.
- **Domain Constants in `DefFileGenerator/def_gen.py`.** Added named constants for Modbus function codes (`MODBUS_COIL`, `MODBUS_DISCRETE`, `MODBUS_HOLDING`, `MODBUS_INPUT`), address boundaries (`MIN_MODBUS_ADDRESS`, `MAX_MODBUS_ADDRESS`), and standard actions (`ACTION_WRITE_ONLY`, `ACTION_READ_ONLY`).

### Fixed

- **Direct Script Execution for Desktop GUI (`DefFileGenerator/gui.py`).** Added parent directory to `sys.path[0]` before package imports, resolving `ModuleNotFoundError: No module named 'DefFileGenerator'` when executing `python DefFileGenerator/gui.py` directly from uninstalled environments. Added headless `--help` and `--version` CLI flags to desktop GUI launcher.
- **Surplus CSV Column Exception Hardening (`DefFileGenerator/extractor.py`).** Resolved `AttributeError: 'list' object has no attribute 'strip'` when extracting CSV files containing empty leading cells and trailing overflow columns (where `csv.DictReader` assigns a list under key `None`). Filtered `None` overflow keys from yielded dictionaries.
- **PyInstaller Virtual Environment Auto-Delegation (`build_exe.py`).** Enhanced `check_pyinstaller` to detect and automatically delegate execution to a local `.venv` containing PyInstaller when executed from an unactivated terminal or base Python interpreter.

### Removed

- **Unused Ad-Hoc Test Artifact (`DefFileGenerator/test_input.csv`).** Removed unreferenced loose CSV file from the core package root directory.
- **Tracked Test Output Artifacts.** Untracked generated test outputs (`stress_test_data/output.csv`, `stress_test_data/output_offset.csv`, `torture_test/output.csv`, `torture_test/output_str.csv`) from version control and updated `.gitignore` with `torture_test/output*.csv` so test batteries execute without dirtying the working directory.

### Changed

- **Core Import Refactoring & Strict Typing (`DefFileGenerator/extractor.py`).** Eliminated dynamic runtime fallback imports (`Generator: Any = None`, `peek_generator: Any = None`, duplicate `_peek_generator_impl`) in favor of direct, statically-typed imports from `DefFileGenerator.def_gen`. Removed redundant `if Generator is not None:` guards in `normalize_type` and `map_and_clean`.
- **Pre-Compiled Regular Expressions.** Converted all 28 type synonym patterns in `normalize_type` to a module-level pre-compiled tuple (`TYPE_SYNONYMS_COMPILED`), eliminating per-invocation regex re-compilation. Pre-compiled address range delimiters, grouping commas, hex address matches, and slug regexes.
- **Constant Time Lookups in Extractor.** Converted table header keywords, PDF metadata markers, communication parameter keywords, and address guards to module-level `frozenset` constants.
- **Python 3.10+ Type Annotations.** Modernized all type annotations to use modern union syntax (`T | None`, `X | Y`) across all modules.
- **Documentation & In-Line Comments.** Added comprehensive Google-style docstrings and educational in-line comments in English across all classes, methods, and functions.

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
  - 4 functional tabs: *Conversion & Generation* (with live register table preview), *Definition Validator* (line-by-line syntax & overlap inspection), *Equipment Templates* (pre-configured Inverters, Meters, Pyranometers, Batteries, Weather Stations), and *Console & Logs*.
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
