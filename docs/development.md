# Developer Guidelines & Engineering Practices

<p align="center">
  <a href="../README.md"><b>README</b></a> •
  <a href="../DEVELOPER_GUIDE.md"><b>Developer Guide</b></a> •
  <a href="architecture.md"><b>Architecture</b></a> •
  <a href="security.md"><b>Security</b></a> •
  <a href="../AUDIT_REPORT.md"><b>Audit Report</b></a>
</p>

---

## 🛠️ Environment Setup

To set up an isolated development environment with all required QA and build tooling:

```bash
# 1. Create and activate a virtual environment
python -m venv .venv
source .venv/bin/activate  # On Windows: .\.venv\Scripts\Activate.ps1

# 2. Install editable package with dev, web, and build extras
pip install --upgrade pip
pip install -e ".[dev,web,build]"
```

---

## 🧪 Quality Assurance Commands

All code modifications must pass the established quality gates with zero errors before pull request submission:

| Check | Tool | Command | Description |
|---|---|---|---|
| **Linting** | Ruff | `ruff check .` | Checks syntax, imports, and static quality rules |
| **Formatting** | Ruff | `ruff format --check .` | Verifies adherence to Black/Ruff styling conventions |
| **Auto-Format** | Ruff | `ruff format .` | Automatically formats all Python code |
| **Type Checking** | Mypy | `mypy DefFileGenerator web` | Strict static type verification |
| **Security Audit** | Bandit | `bandit -r DefFileGenerator web -ll` | Static security scanner for high/medium vulnerabilities |
| **Test Suite** | Pytest | `pytest` | Runs the full 685+ automated test suite |
| **Test Coverage** | Pytest-Cov | `pytest --cov=DefFileGenerator --cov=web` | Generates detailed line/branch coverage report |
| **Wheel Build** | Build / UV | `uv build` (or `python -m build`) | Verifies packaging contract and asset inclusions |

---

## 🛡️ Testing Strategy

- **Deterministic & Network-Free**: All unit tests must execute deterministically without requiring internet or external server access.
- **End-to-End Pipelines**: Integration tests must cover the complete flow: document extraction $\rightarrow$ heuristic mapping $\rightarrow$ data normalization $\rightarrow$ overlap checks $\rightarrow$ atomic writing.
- **Regression Invariance**: Every bug fix must introduce a minimal reproducing test case in [`DefFileGenerator/tests/`](../DefFileGenerator/tests/).
- **Stress & Benchmark Testing**: Benchmark suites evaluate performance on large register maps (5,000+ registers) using [`DefFileGenerator/tests/stress_test_gen.py`](../DefFileGenerator/tests/stress_test_gen.py).
- **Dead-Code Auditing**: Run `vulture` periodically to detect dead code and unused imports:
  ```bash
  vulture DefFileGenerator web build_exe.py generate_webdyn_def.py --min-confidence 80
  ```

---

## 📜 Contribution Rules

1. **Single Responsibility**: Each pull request must address a single, well-defined bug fix, feature, or refactoring task.
2. **Backward Compatibility**: Preserve existing CLI flags, public programmatic APIs (`GeneratorConfig`, `generate_webdyn_definition`), and official WebdynSunPM CSV layouts.
3. **Memory Efficiency**: Always favor generator pipelines (`yield`) over large in-memory list allocations for rows and records. Never return a generator attached to an already closed file resource.
4. **Code Cleanliness**: Maintain zero warnings under `ruff` and `mypy`. Avoid committing transient caches (`.pytest_cache`, `.mypy_cache`, `dist/`, `build/`).

---

<p align="center">
  <a href="../README.md"><b>⬅ Back to README</b></a> •
  <a href="security.md"><b>Security Model ➡</b></a>
</p>
