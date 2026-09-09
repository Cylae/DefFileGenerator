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
|                       Two-Pass Heuristic Mapping                        |
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

1. **Lazy Generator Pipelines**: File inputs are parsed line-by-line using Python generator iterators (`yield`), ensuring $O(1)$ memory retention even when parsing 10,000+ registers.
2. **Determinism**: Outputs are guaranteed to be byte-identical across identical inputs, unaffected by `PYTHONHASHSEED`.
3. **Decoupled Responsibilities**: Extractor deals strictly with document parsing and column mapping heuristics; Generator handles data-type logic, address offsets, and overlap validations.
4. **Defensive Parsing**: Hostile inputs (malformed XML entities, CSV formula injection, control characters) are handled at the system boundary before reaching processing state.
