# Software Architecture & Engineering Overview

The **DefFileGenerator** project is built as a high-performance, modular Python engine designed to parse industrial Modbus register documentation in various unstructured formats (PDF, Excel, CSV, XML) and emit validated, structured WebdynSunPM definition files (`.csv`).

---

## 🏛️ Component Architecture

```text
+-------------------------------------------------------------------------+
|                              CLI / API Entry                            |
|             (DefFileGenerator/main.py / generate_webdyn_def.py)          |
+-------------------------------------------------------------------------+
                                    |
                                    v
+-------------------------------------------------------------------------+
|                            Extractor Module                             |
|                    (DefFileGenerator/extractor.py)                      |
|                                                                         |
|  +----------------+  +----------------+  +--------------+  +---------+  |
|  | pdfplumber     |  | openpyxl       |  | csv.Sniffer  |  | defused |  |
|  | (PDF Tables)   |  | (Excel Sheets) |  | (CSV Stream) |  | (XML)   |  |
|  +----------------+  +----------------+  +--------------+  +---------+  |
|                                   |                                     |
|          Exact -> synonym -> partial -> PDF glyph fallback mapping     |
|                     & Row Dict Yielding Generators                      |
+-------------------------------------------------------------------------+
                                    |
                                    v
+-------------------------------------------------------------------------+
|                            Generator Module                             |
|                    (DefFileGenerator/def_gen.py)                        |
|                                                                         |
|  - Address Normalization (Hex, Decimal, Offset Shifts)                  |
|  - Data Type Standardization (U16, I32, F32, Endian Suffixes)           |
|  - O(log N) Bisect Register Overlap Validation                         |
|  - Linear Coefficient Scaling (CoefA = Factor * 10^Scale, CoefB = Offset) |
|  - CSV Injection Escaping & Control Character Stripping                 |
+-------------------------------------------------------------------------+
                                    |
                                    v
+-------------------------------------------------------------------------+
|                           Output Writer & FS                            |
|                 (Atomic Staging & UTF-8-SIG CSV Formatting)             |
+-------------------------------------------------------------------------+
```

---

## ⚡ Key Architectural Invariants

1. **Streaming-Oriented Pipelines**: CSV/XML and mapped rows use iterators to limit retention. Some formats and boundaries necessarily materialize data: PDF/Excel libraries expose document structures, overlap validation retains intervals, and the web API caps and materializes at most 65,536 mapped registers.
2. **Determinism**: Outputs are guaranteed to be byte-identical across identical inputs, unaffected by `PYTHONHASHSEED`.
3. **Decoupled Responsibilities**: Extractor deals strictly with document parsing and column mapping heuristics; Generator handles data-type logic, address offsets, and overlap validations.
4. **Defensive Parsing**: Hostile inputs (malformed XML entities, CSV formula injection, control characters) are handled at the system boundary before reaching processing state.
5. **Conservative Completeness**: Extraction never fabricates missing addresses. Image-only PDFs require OCR, while non-tabular PDF geometry may require preprocessing or a future geometry-aware reader.

## Public surfaces

- `DefFileGenerator/main.py`: canonical `deffilegen` CLI (`run`, `extract`, `generate`, `validate`, `gui`).
- `DefFileGenerator/gui.py`: desktop entry point exposed as `deffilegen-gui`.
- `web/app.py`: FastAPI application and bundled static frontend.
- `generate_webdyn_def.py`: supported programmatic compatibility API.
- `doc_to_webdyn.py`: legacy one-command wrapper retained for backward compatibility; new documentation uses `deffilegen run`.
