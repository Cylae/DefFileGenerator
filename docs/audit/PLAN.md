# Apollo-Standard Engineering Plan: DefFileGenerator

## 1. Starting State & Scope (Captured 2026-11-25)

- **Repository**: `https://github.com/Cylae/DefFileGenerator.git`
- **Local Workspace**: `c:\Users\alexandre.calais\Documents\GitHub\DefFileGenerator`
- **Starting Commit / Branch**: `main` at `346e4546` (clean working tree at start; 7 commits ahead of `origin/main`)
- **Runtime & Toolchain**: Python 3.15.0 (`.venv`), `pytest` 9.1.1, `ruff` 0.17.0, `mypy` 2.4.0, `bandit` 1.9.4, `pyinstaller` 6.22.3, Node.js
- **Shared Storage Constraint**: Never write/modify/delete on `\\WDN-FILES.WEBDYN.COM` without explicit user authorization.

### Two Separate Workstreams
1. **PHASE A — EXECUTE NOW**: Full-project audit, root-cause defect remediation, justified refactoring, falsifiable regression tests, packaging/CI verification, and ledger/doc synchronization across the entire first-party repository.
2. **PHASE B — PREPARE NOW (DESIGN-ONLY)**: Produce/upgrade an implementation-ready specification, gap/reusability analysis, component contracts, acceptance policy, caching/versioning model, and executable test plan in `docs/AUTONOMOUS_MODBUS_DISCOVERY_SPECIFICATION.md` for autonomous online Modbus discovery. Do NOT enable live online search providers, site crawlers, or LLM extractors in Phase A.

---

## 2. Baseline Measurements at `346e4546`

- `pytest -q`: **4,701 collected** -> **4,693 passed, 3 failed, 5 skipped** (`test_health_check`, `test_frontend_validity_download_and_stale_request_contracts`, `test_mypy_typecheck_passes`).
- `ruff check .`: **43 errors** (`EXECUTED_FAIL`).
- `ruff format --check .`: **18 unformatted files** (`EXECUTED_FAIL`).
- `mypy DefFileGenerator web`: **1 error** (`web\app.py:1079`, `EXECUTED_FAIL`).
- `bandit -r DefFileGenerator web -ll`: **0 Medium/High issues** (`EXECUTED_PASS`).
- `node DefFileGenerator/tests/interface_frontend_contracts.js web/static/app.js`: **ReferenceError (`localStorage`)** (`EXECUTED_FAIL`).
- `python scripts/validate_production.py`: **1,194 / 1,194 passed** (`EXECUTED_PASS`).

---

## 3. Execution Milestones

### Milestone A.1: Full Reconnaissance, Baseline Capture & Audit Ledgers [IN_PROGRESS]
- Read 100% of all first-party Python, Web, JS/HTML/CSS, build, and script files via bounded chunks across 8 parallel explorers.
- Establish `FEATURE_INVENTORY.md`, `REVIEW_MATRIX.md`, `FINDINGS.md` (`FIND-030`+), `PLAN.md`, `VALIDATION.md`, and `PERFORMANCE.md`.

### Milestone A.2: Baseline CI & Core Pipeline Hardening
- Fix immediate baseline failures (`web/app.py:1079` mypy, `app.js`/`interface_frontend_contracts.js` `localStorage`, `test_health_check` contract, all ruff lint/format errors).
- Harden Stream 1 (`extractor.py`, `pdf_parser.py`, `excel_modbus_extractor.py`, `batch_parser.py`): bomb/size guards, handle closure, header/radix/scaling fixes, parallel batch safety.
- Harden Stream 2 (`def_gen.py`, `knowledge_miner.py`, `production_governor.py`, `report_generator.py`, `duplicate_detector.py`, `version_manager.py`, `io_utils.py`): radix `0010`, `bool`-as-`int`, `strict_crlf` mixed-LF bypass, tag/index repairs, atomic persistence, `commit_id` traversal, diff FC keying.

### Milestone A.3: Interfaces, Existing Discovery Scaffold, SunSpec & Generators
- Harden Stream 3 (`main.py`, `gui.py`, `DefFileGenerator/discovery/*`): CLI flags/guards/exit codes, non-blocking GUI sheet inspection/validation, SSRF DNS pinning + non-fabrication in `SourceDiscovery` + `WEBDYN_CLASSIC` export + real category/filename propagation.
- Harden Stream 4 (`sunspec/*`, `dyness_generator.py`, `mars_generator.py`, `x7_generator.py`): SunSpec CSV injection, SF sentinel `-32768`, PICS `200`-row cap & handle leaks, Model 123 SF ref & registry fixes, X7 duplicate tags & `STRING_24` false positive, atomic staged writes across all domain generators.

### Milestone A.4: Web Studio, Vendor Engines & Packaging/Scripts
- Harden Stream 5 (`web/app.py`, `web/static/*`, `web/templates/index.html`, `interface_frontend_contracts.js`): fix broken Discovery UI contract, path confinement on `versions`/`duplicates`/`library`, CSV injection in `/api/definition/revalidate`, `inspectModal` >1000-row loss, async offloading, XSS escaping, CORS tightening.
- Refactor & harden Stream 6 (`*_definitions_engine.py`, `DefFileGenerator.spec`, `build_exe.*`, `launch_gui.bat`, `generate_webdyn_def.py`, `doc_to_webdyn.py`, `scripts/*`): extract shared engine base module to eliminate 7x duplication, atomic staged writes, fix `Sungrow` `.xlsx`->CSV bug, `sma` case-sensitive `Sources/` bug, `.spec` hiddenimports, `.bat` `!ERRORLEVEL!` and exit-code propagation, `scripts/` error reporting & ruff cleanup.

### Milestone A.5: Falsifiable Verification, Benchmarks, Wheel & Clean-Install
- Add falsifiable regression/fuzz/boundary tests for every fixed defect (`FIND-030`+).
- Re-benchmark runtime & memory vs baseline and run soak check; verify wheel build and clean out-of-tree CLI/import execution.

### Milestone B.0: Phase B Design-Only Specification Upgrade
- Update `docs/AUTONOMOUS_MODBUS_DISCOVERY_SPECIFICATION.md` covering all Sections 11–27 (exact reuse analysis against audited `DefFileGenerator/discovery/*` & core modules, missing concrete interfaces, equipment identity/disambiguation, source ranking, retrieval SSRF/budgets, OCR/multi-format parsing, canonical IR, Modbus & SunSpec semantics, WebdynSunPM & Classic target profiles, completeness reconciliation, multi-dimensional confidence & discrete outcomes, 3-tier cache & versioning, optional bounded LLM policy, full offline/live test plan, milestones, and concrete unresolved product decisions).

### Milestone A.6: Dual Independent Peer Reviews & Frozen-State Gatekeeper Verification
- Run fresh-context Red-Team Failure Audit and Regression/Diff Audit (`reviewer` subagent) + full frozen-state validation (`ruff`, `mypy`, `bandit`, `pytest --cov`, `node`, `validate_production`, wheel check).
