# Apollo-Standard Performance & Efficiency Report: DefFileGenerator

## 1. Baseline Measurements (Commit `346e4546`, Pre-Remediation)

- **Environment**: Windows 11 x86_64, Python 3.15.0 (`.venv`), `tracemalloc` active, root `WARNING` logs suppressed to measure I/O and compute rather than synchronous console stdout flush.
- **Workloads**:

| Workload | Scope | Baseline Time | Baseline Peak Traced Mem | Notes |
|---|---|---:|---:|---|
| **5-Pillars Production Scan** | 1,194 CSV files in `production/` via `verify_file_pillars` | 27.316 s | 21.59 MB (combined run) | Each CSV file opened/read 3 times in `verify_file_pillars` (`read_bytes` + `csv.reader` + `validate_csv_detailed`) |
| **Duplicate Detector Scan** | 1,194 CSV files in `production/` via `DuplicateDetector.scan` | 5.044 s | — | Reads every file twice (`read_bytes` + `read_text`) |
| **5,000-Row CSV Generation** | 5,000 synthetic `U16` rows via `run_generator(..., strict_validation=True)` | 1.028 s | — | Includes `extract_from_csv` + `process_rows` + staged write + `validate_csv_detailed` |
| **X7 Workbook Extraction** | `X7-ModbusTcp-PORT502-20261105-V1.0.6-Beta.xlsx` (428 mapped rows) | 4.372 s | — | Uses `openpyxl.load_workbook(..., data_only=True)` + secondary raw formula pass |

---

## 2. Post-Remediation Comparisons & Soak Tests

*To be measured after Phase A fixes are integrated.*
