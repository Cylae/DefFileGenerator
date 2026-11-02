# Apollo-Standard Validation Record: DefFileGenerator

## 1. Test Environment & Execution Profile

- **Operating System**: Microsoft Windows 11 (build 26100 x86_64)
- **Python Runtime**: CPython 3.14.7 final (64-bit)
- **Package & Dependency Manager**: `uv` 0.12.9
- **Test Framework**: `pytest` 9.1.1, `pytest-cov` 7.1.0, `pluggy` 1.6.0
- **Code Analyzers**: `ruff` 0.6+, `mypy` 1.10+, `bandit` 1.7+

---

## 2. Gate Verification Summary

| Gate | Command | Scope | Result | Observed Evidence |
| :--- | :--- | :--- | :--- | :--- |
| **Linter** | `uv run ruff check .` | Entire repository | **EXECUTED_PASS** | All checks passed (0 errors across all files) |
| **Formatter** | `uv run ruff format --check .` | Entire repository | **EXECUTED_PASS** | 129 files already formatted cleanly |
| **Type Checker** | `uv run mypy DefFileGenerator web` | First-party packages | **EXECUTED_PASS** | Success: 0 issues found in 97 source files |
| **Security Audit** | `uv run bandit -r DefFileGenerator web -ll` | First-party packages | **EXECUTED_PASS** | 0 high/medium issues across 22,583 LOC |
| **Build & Packaging** | `uv build --offline` | Sdist & Wheel | **EXECUTED_PASS** | `def_file_generator-0.2.1-py3-none-any.whl` generated |
| **Package Integrity** | Python ZipFile inspection | Built wheel archive | **EXECUTED_PASS** | All 8 required web/core files present; 0 test leaks |
| **CLI Verification** | `uv run deffilegen --version` | Application CLI | **EXECUTED_PASS** | `deffilegen 0.2.1` returned with exit code 0 |
| **Unit & Integration** | `uv run pytest --cov=DefFileGenerator --cov=web` | 1,040 collected tests | **EXECUTED_PASS** | 1,039 passed, 1 skipped, 0 failed |

---

## 3. Detailed Coverage Metrics

| Subsystem | Statements | Missing | Branch | Partial | Coverage | Status |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| `DefFileGenerator/__init__.py` | 6 | 0 | 0 | 0 | **100%** | Verified |
| `DefFileGenerator/def_gen.py` | 921 | 26 | 386 | 17 | **97%** | Verified |
| `DefFileGenerator/extractor.py` | 849 | 63 | 454 | 38 | **92%** | Verified |
| `DefFileGenerator/knowledge_miner.py` | 618 | 38 | 386 | 34 | **92%** | Verified |
| `DefFileGenerator/report_generator.py` | 216 | 8 | 40 | 7 | **93%** | Verified |
| `DefFileGenerator/mars_generator.py` | 335 | 27 | 176 | 25 | **89%** | Verified |
| `DefFileGenerator/io_utils.py` | 41 | 0 | 10 | 0 | **100%** | Verified |
| `web/app.py` | 205 | 9 | 64 | 10 | **93%** | Verified |
| `DefFileGenerator/main.py` | 436 | 78 | 164 | 25 | **78%** | Verified |
| `DefFileGenerator/gui.py` | 823 | 176 | 230 | 50 | **72%** | Verified |
| **TOTAL** | **4,450** | **425** | **1,910** | **206** | **88%** | **PASS** |

---

## 4. Skip Rationale

- `DefFileGenerator/tests/test_coverage_wave28.py:213` (`test_real_world_definitions_audit_from_share`):
  - **Reason**: Conditioned on live Webdyn internal network share `\\WDN-FILES.WEBDYN.COM\Support\Equipementiers`.
  - **Classification**: `NOT_APPLICABLE` in offline/local environments without active corporate VPN; guarded with fast zero-latency check.

