# Apollo-Standard Finding Ledger: DefFileGenerator

| Finding ID | Severity | Category | Affected Paths | Summary | Status |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **FIND-001** | MEDIUM | Performance / Network | `DefFileGenerator/tests/test_50_equipementiers.py` | UNC path probe causes 30-45s SMB network timeout outside Webdyn corporate LAN | OPEN |
| **FIND-002** | LOW | Quality / Linter | `scratch/`, `pyproject.toml` | 40 ruff linter issues in tracked `scratch/` scripts | OPEN |
| **FIND-003** | MEDIUM | Portability / Sandbox | `web/app.py`, `starlette.testclient` | Asyncio loopback socket hang in AppContainer sandbox on Windows | OPEN |

---

## Detailed Findings

### FIND-001: UNC Path Probe Induces Blocking SMB Network Timeout
- **Severity**: MEDIUM
- **Component**: Test Suite (`test_50_equipementiers.py`)
- **Trigger**: Executing `test_50_battery` without `EQUIPEMENTIERS_DIR` environment variable set.
- **Root Cause**: `os.path.exists(r"\\WDN-FILES.WEBDYN.COM\Support\Equipementiers")` triggers Windows SMB redirector name resolution and TCP 445 connection attempts, which stall for 30–45s before returning `False`.
- **Expected Behavior**: Fast-fail detection of UNC network path without stalling test execution.
- **Remediation**: Probe network host connectivity with a short socket timeout (e.g., 500ms) before invoking `os.path.exists` on UNC paths.

### FIND-002: Ruff Lint Violations in Tracked Scratch Scripts
- **Severity**: LOW
- **Component**: Build & Tooling (`pyproject.toml`, `scratch/*.py`)
- **Trigger**: Running `ruff check .` across entire repository.
- **Root Cause**: Scripts committed to `scratch/` contained unformatted imports, unused variables, and unnecessary string formatting.
- **Expected Behavior**: `ruff check .` passes with 0 errors across all repository files.
- **Remediation**: Clean up scratch scripts with `ruff format` and `ruff check --fix`, and add appropriate exclusions in `pyproject.toml` if scratch files are intended as transient exploratory scripts.

### FIND-003: Starlette / AnyIO TestClient Hangs in Loopback-Restricted Windows Sandbox
- **Severity**: MEDIUM
- **Component**: Web API Testing (`web/app.py`, `tests/test_web*.py`)
- **Trigger**: Running FastAPI `TestClient` tests in a Windows sandbox where network loopback (`127.0.0.1`) is restricted.
- **Root Cause**: On Windows, Python 3.14 `asyncio.ProactorEventLoop._make_self_pipe` creates an internal loopback socket via `_fallback_socketpair`. In a strictly network-isolated environment without loopback exemption, `accept()` stalls indefinitely.
- **Expected Behavior**: Clean diagnostics or non-blocking behavior when network loopback is unavailable.
- **Remediation**: Document the requirement for Windows loopback socket permissions, and provide fallback or explicit test tagging for web endpoints requiring loopback transport.


### FIND-004: Version string hardcoded in five places and drifted from pyproject (0.2.1 vs 0.3.0)
- **Severity**: MEDIUM
- **Component**: `DefFileGenerator/main.py`, `gui.py`, `web/app.py` (OpenAPI + `/api/health`), tests
- **Trigger**: `deffilegen --version`, `/api/health`, OpenAPI docs reported 0.2.1 while the package is 0.3.0; baseline `test_package_metadata` failed.
- **Root Cause**: No single source of truth for the version.
- **Fix**: `DefFileGenerator/_version.py` is the only definition; CLI, GUI, web and `__init__` import it. Tests derive the expected value from it.
- **Regression**: `tests/test_version_consistency.py` (pyproject == package == web; fails if a stale literal returns; verified by reverting `web/app.py`).
- **Status**: RESOLVED

### FIND-005: Mars definitions test regenerated committed fixtures in place
- **Severity**: LOW
- **Component**: `tests/test_mars_requested_definitions.py`
- **Trigger**: Running the suite rewrote 10 tracked `generated_definitions/*.csv` (LF in git, CRLF from the generator), dirtying the working tree.
- **Fix**: Regenerate into a `TemporaryDirectory`; fall back to committed files only when the source workbook is absent.
- **Status**: RESOLVED (open follow-up: no `.gitattributes`, so committed LF vs generated CRLF is undecided; needs a product decision on the canonical line ending)

### FIND-006: Vacuous batch skip-existing test leaked output into the repo
- **Severity**: LOW
- **Component**: `tests/test_coverage_wave31_cli_and_reports.py::test_batch_command_skip_existing_and_exceptions`
- **Root Cause**: Namespace used `output=` but `batch_command` reads `output_dir`, so the pre-created file was never consulted and `generated_definitions/huawei_sun2000_definition.csv` was written to the CWD.
- **Fix**: Use `output_dir=`; the skip-existing assertion now actually exercises the code path.
- **Status**: RESOLVED
