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

