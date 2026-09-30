# Conversion reliability and project verification

This enhancement focuses on preserving input documents and existing outputs, reporting conversion failures consistently, and making generated definitions agree across the CLI, desktop GUI, public Python API, and web application.

## Changes

- Generated CSV files are staged and validated before publication when validation is configured. Failed extraction, invalid definitions, and empty results preserve an existing destination. Configured stdout and caller-owned streams are validated before copying, and remain open.
- Source/destination aliases are rejected, including existing hard links and symbolic links. Templates and intermediate extraction files use staged writes. Intermediate CSV values receive spreadsheet-formula protection.
- Address validation checks the complete occupied register span. Strict overlap checks normalize register groups. Byte and word swap aliases, zero addresses, explicit zero factors, gain precision, headerless Excel rows, RAW lengths, and PDF continuation records are corrected.
- Truncated or malformed declared Unicode, NUL-containing CSV, partial parser failures, XML DTDs, oversized Excel archives, and excessive worksheet dimensions produce explicit failures rather than silently incomplete success.
- Web uploads are limited before multipart parsing, including requests without Content-Length. File counts and mapped register counts are bounded. Blocking conversion runs outside the async event loop. Previews reflect the actual generated CSV, and browser results are invalidated when inputs change.
- The wheel includes both documented Python API helpers, the shared I/O helper, and web assets. Windows release archives include the MIT license. Documentation describes current behavior and labels the old audit report as historical.

## Local verification

| Check | Result |
| --- | --- |
| Python 3.10.21 | 794 passed, 9 skipped; 35 subtests passed |
| Python 3.11.16 | 794 passed, 9 skipped; 35 subtests passed |
| Python 3.12.10 | 802 passed, 1 skipped; 35 subtests passed |
| Runtime coverage on Python 3.12 | 87%, including branch coverage; excludes bundled test modules |
| Ruff lint and formatting | Passed |
| mypy | Passed for 67 source files |
| Bandit at the configured medium/high threshold | No medium or high findings |
| Frozen dependency lock | Passed |
| Installed environment dependency audit | No known vulnerabilities reported after updating the isolated build environment's pip |
| Wheel and source distribution | Built successfully |
| Windows CLI and GUI executables | Built successfully; CLI version smoke test passed |

The single skipped test needs the external `Equipementiers` document collection. Python 3.10 and 3.11 additionally skip eight GUI tests because their locally managed runtimes lack usable Tcl data; the Python 3.12 GUI tests pass. GitHub CI supplies independent Linux verification on the three supported Python versions.

## Impact and limits

The changes affect every conversion entry point and the distributable package. Inputs that previously produced misleading partial success may now fail explicitly. Strict mode can reject malformed or overlapping definitions that formerly reached a destination; the diagnostic CLI opt-out is documented.

An arbitrary caller-owned stream cannot be rolled back if the stream itself fails while copying an already validated result. Staged file replacement depends on the destination filesystem; a killed process can leave a temporary file. PDF extraction still depends on document layout and needs human review before deploying a definition to equipment. These checks demonstrate the tested behavior, not a guarantee that every possible manufacturer document or defect has been covered.
