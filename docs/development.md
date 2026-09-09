# Developer Guidelines & Engineering Practices

## Environment Setup

```bash
pip install -e ".[dev]"
```

## Quality Assurance Commands

```bash
# Code linting
ruff check .

# Code formatting check
ruff format --check .

# Static type checking
mypy DefFileGenerator generate_webdyn_def.py doc_to_webdyn.py

# Full test suite execution
pytest

# Test coverage reporting
pytest --cov=DefFileGenerator --cov-report=term-missing
```

## Testing Strategy

- **Unit Tests**: Must be deterministic, fast, isolated, and require zero network access.
- **Integration Tests**: Verify the end-to-end pipeline: extraction → field mapping → generation → overlap validation.
- **Regression Protection**: Every bug fix must include a minimal reproducing test case in `DefFileGenerator/tests/`.
- **Stress & Battery Tests**: Large-scale benchmark datasets (e.g. 5,000+ registers) are generated on demand via `DefFileGenerator/tests/stress_test_gen.py` and evaluated via `run_gigantic_battery.py`.

## Contribution Rules

1. **Single Responsibility**: Pull requests should address one clear feature or fix.
2. **Backward Compatibility**: Preserve existing CLI arguments, public Python APIs (`GeneratorConfig`, `generate_webdyn_definition`), and Webdyn CSV output structures.
3. **Resource Efficiency**: Use lazy generator pipelines (`yield`) for processing row iterables to keep memory usage $O(1)$. Never return a generator bound to an already closed file resource.
4. **Clean Codebase**: Maintain clean, fully typed Python 3.10+ code with zero `ruff` or `mypy` warnings.

## Performance Metrics

Always measure execution times before and after making structural changes. Critical benchmarks include total execution time, row streaming throughput, peak memory usage, and type normalization latency.
