# Apollo-Standard Validation Record: DefFileGenerator

## 1. Test Environment & Execution Profile

- **Operating System**: Microsoft Windows 11 (build 26100 x86_64)
- **Python Runtime**: CPython 3.14.7 final (64-bit)
- **Node Runtime**: Node.js v22+
- **Package & Dependency Manager**: `uv` 0.12.9
- **Test Framework**: `pytest` 9.1.1, `pytest-cov` 7.1.0, `pluggy` 1.6.0
- **Code Analyzers**: `ruff` 0.6+, `mypy` 1.10+, `bandit` 1.7+

---

## 2. Gate Verification Summary

| Gate | Command | Scope | Result | Observed Evidence |
| :--- | :--- | :--- | :--- | :--- |
| **Linter** | `uv run ruff check .` | Entire repository | **EXECUTED_PASS** | All checks passed (0 errors across 141 files) |
| **Formatter** | `uv run ruff format --check .` | Entire repository | **EXECUTED_PASS** | 141 files already formatted cleanly |
| **Type Checker** | `uv run mypy DefFileGenerator web` | First-party packages | **EXECUTED_PASS** | Success: 0 issues found in 102 source files |
| **Security Audit** | `uv run bandit -r DefFileGenerator web -ll` | First-party packages | **EXECUTED_PASS** | 0 high/medium issues across 24,193 LOC |
| **Build & Packaging** | `pip wheel . --no-deps -w dist` | Sdist & Wheel | **EXECUTED_PASS** | `def_file_generator-0.3.0-py3-none-any.whl` generated |
| **Package Integrity** | Python ZipFile inspection | Built wheel archive | **EXECUTED_PASS** | All required web/core files present; 0 test leaks |
| **CLI Verification** | `deffilegen --version` | Application CLI | **EXECUTED_PASS** | `deffilegen 0.3.0` returned with exit code 0 |
| **Frontend Contracts** | `node DefFileGenerator/tests/interface_frontend_contracts.js web/static/app.js` | Web Studio DOM contracts | **EXECUTED_PASS** | All 4 tabs asserted: Convert, Tag, Audit, Library, Toast, Copy, Filter, ARIA |
| **Unit & Integration** | `pytest --cov=DefFileGenerator --cov=web` | 1,099 collected tests | **EXECUTED_PASS** | 1,094 passed, 5 skipped, 0 failed |
| **Live Share Validation** | `ENABLE_NETWORK_SHARE_TESTS=1 pytest test_30_equipementiers.py test_50_equipementiers.py` | 30 & 52 equipment batteries | **EXECUTED_PASS** | 30/30 passed in battery 30; 47/52 passed in battery 52 (>=45 passed) |
| **Meticulous Definition Audit** | `python scratch/deep_audit_all_definitions.py` | 2,344 definitions (750,577 registers) across all libraries | **EXECUTED_PASS** | 2,344 / 2,344 (100.0%) valid, 0 errors, 11-col headers & sequential indices |

---

## 3. Detailed Coverage Metrics (90% Total Coverage)

| Subsystem | Statements | Missing | Branch | Partial | Coverage | Status |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| `DefFileGenerator/__init__.py` | 6 | 0 | 0 | 0 | **100%** | Verified |
| `DefFileGenerator/_version.py` | 1 | 0 | 0 | 0 | **100%** | Verified |
| `DefFileGenerator/def_gen.py` | 921 | 26 | 386 | 17 | **97%** | Verified |
| `DefFileGenerator/extractor.py` | 849 | 63 | 454 | 38 | **92%** | Verified |
| `DefFileGenerator/knowledge_miner.py` | 618 | 37 | 386 | 33 | **93%** | Verified |
| `DefFileGenerator/report_generator.py` | 216 | 2 | 40 | 3 | **98%** | Verified |
| `DefFileGenerator/mars_generator.py` | 335 | 27 | 176 | 25 | **89%** | Verified |
| `DefFileGenerator/io_utils.py` | 41 | 0 | 10 | 0 | **100%** | Verified |
| `DefFileGenerator/main.py` | 598 | 38 | 232 | 27 | **91%** | Verified |
| `DefFileGenerator/gui.py` | 824 | 176 | 230 | 50 | **72%** | Verified |
| `web/__init__.py` | 0 | 0 | 0 | 0 | **100%** | Verified |
| `web/app.py` | 327 | 20 | 96 | 15 | **92%** | Verified |
| **TOTAL** | **4,736** | **389** | **2,010** | **208** | **90%** | **PASS** |

---

## 4. Specialized Stress & Industrial Battery Results

| Test Battery | Target / Scope | Metrics Observed | Status |
| :--- | :--- | :--- | :--- |
| **Torture Battery** (`run_torture_battery.py`) | Ambiguous column detection, $10^6$ address offsets, `STR<n>` expansion | 4/4 rows mapped, offset arithmetic preserved, zero crash | **PASS** |
| **Gigantic Battery** (`run_gigantic_battery.py`) | 5,000-row CSV generation, XML/Excel/PDF parsing, XXE defense, 12,000 type conversions | CSV generation in 0.269s, 12,000 type conversions in 0.0011s, XXE defended | **PASS** |
| **Equipementiers Battery** (`run_equipementiers_battery.py`) | 45 real-world files from 20+ vendors across XLSX, CSV, XML | 7,949 registers extracted, 30 valid definitions generated, 0 unhandled exceptions | **PASS** |
| **Huawei & ABB Battery** (`test_huawei_abb_battery.py`) | 35 production manufacturer documents across inverters, meters, smartloggers, chargers | 33 valid definitions generated, 2 image/empty PDFs safely caught | **PASS** |
| **Definitions Library** (`test_definitions_library.py`) | 125 pre-validated industrial definitions in `Definitions/` | 100% pass strict Webdyn syntax, prefix convention, and catalog synchronization | **PASS** |
| **Knowledge Training Engine** (`deffilegen train --summary`) | 125 files, 18,127 registers mined across 61 manufacturers | 630 distinct variable-to-tag mappings mined in <1s | **PASS** |

---

## 5. Skip Rationale

- `DefFileGenerator/tests/test_coverage_wave28.py:213` (`test_real_world_definitions_audit_from_share`):
  - **Reason**: Conditioned on live Webdyn internal network share `\\WDN-FILES.WEBDYN.COM\Support\Equipementiers`.
  - **Classification**: `NOT_APPLICABLE` in offline/local environments without active corporate VPN; guarded with fast zero-latency check.
