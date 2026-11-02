# Architecture & Engineering Overview

<p align="center">
  <a href="../README.md"><b>README</b></a> •
  <a href="../DEVELOPER_GUIDE.md"><b>Developer Guide</b></a> •
  <a href="input-format.md"><b>Input Formats</b></a> •
  <a href="security.md"><b>Security</b></a> •
  <a href="development.md"><b>Development</b></a> •
  <a href="audit/VALIDATION.md"><b>Audit Ledger</b></a>
</p>

---

The **DefFileGenerator** project is engineered as a high-performance, modular Python engine designed to parse industrial Modbus register documentation in heterogeneous unstructured formats (**PDF, Excel, CSV, XML**) and produce validated, structured WebdynSunPM definition files (`.csv`).

---

## 🏛️ System Architecture Diagram

```mermaid
flowchart TD
    subgraph Input["📥 Heterogeneous Source Documents"]
        A1["📄 PDF (pdfplumber Tables)"]
        A2["📊 Excel (openpyxl read_only)"]
        A3["📑 CSV (Sniffer / Multi-Encoding)"]
        A4["🏷️ XML (defusedxml anti-XXE)"]
    end

    subgraph Extractor["⚙️ Extraction Module (extractor.py)"]
        B1["Smart Header Detection & Keyword Scoring"]
        B2["Multi-Row Header Banner Merging"]
        B3["Tiered Column Mapping (Exact > Synonym > Partial > PDF Glyphs)"]
        B4["Streaming Row Dict Generators"]
    end

    subgraph Semantic["🧠 Semantic Knowledge Miner (knowledge_miner.py)"]
        S1["211,000+ Register Industrial Training Dataset"]
        S2["Canonical SunPM Tag Attribution (P_AC, E_TOT, U_AC_1...)"]
        S3["Multi-Signal Disambiguation (Phase, Absolute vs %, S vs Q)"]
        S4["Action Code Inference (4=Telemetry, 9=Alarm, 1=Command)"]
    end

    subgraph Generator["🔧 Normalization Engine (def_gen.py)"]
        C1["Address Normalization (Hex 0x..., Decimal, Ranges)"]
        C2["Data Type Standardization (U16, I32, F32, BITS, Endian)"]
        C3["Linear Scaling Coefficients (CoefA, CoefB)"]
        C4["Unique Tag Synthesis & Deduplication"]
    end

    subgraph Validation["🛡️ Validation & Security Layer"]
        D1["Bisect Interval Address Overlap Check"]
        D2["Bit-Slice Constraints (startbit 0..15, length 1..16)"]
        D3["CSV Formula Injection Escaping"]
        D4["Memory Boundary Checks (0..65535)"]
    end

    subgraph Output["💾 Atomic Persistence & Public Surfaces"]
        E1["Staged Sibling File (.tmp)"]
        E2["fsync() + os.replace() Atomic Rename"]
        E3["Validated WebdynSunPM CSV (UTF-8 BOM)"]
        E4["Technical Quality Reports (report_generator.py)"]
    end

    Input --> Extractor
    Extractor --> Semantic
    Semantic --> Generator
    Generator --> Validation
    Validation --> Output
```

---

## 🏗️ Detailed Component Data Flow

```text
[ PDF / Excel / CSV / XML ]
            │
            ▼
┌────────────────────────────────────────────────────────┐
│  DefFileGenerator/extractor.py                         │
│  - Multi-page table extraction via pdfplumber          │
│  - Stream buffer ingestion (avoids OS file locks)      │
│  - Smart header detection & Modbus keyword scoring     │
│  - Multi-row header banner merging                     │
│  - Exact -> synonym -> partial -> PDF glyph fallback   │
└───────────────────────────┬────────────────────────────┘
                            │ Yields intermediate register dicts
                            ▼
┌────────────────────────────────────────────────────────┐
│  DefFileGenerator/knowledge_miner.py                   │
│  - 211,000+ register training knowledge base           │
│  - Official SunPM legacy-to-canonical tag mapping     │
│  - Phase detection (_1, _2, _3) and unit disambiguation │
│  - Action code inference (4=Telemetry, 9=Alarm, 1=Cmd) │
└───────────────────────────┬────────────────────────────┘
                            │ Enriched register stream
                            ▼
┌────────────────────────────────────────────────────────┐
│  DefFileGenerator/def_gen.py                           │
│  - Address Normalization (Hex 0x..., Decimal, Ranges)  │
│  - Data Type Standardization (U16, I32, F32, Endian)   │
│  - BITS compound slicing (address_startbit_length)     │
│  - Bisect Interval Lookup and Overlap Detection       │
│  - Linear Scaling (CoefA = Factor * 10^Scale, CoefB)   │
│  - CSV Injection Escaping & Control Character Stripping │
└───────────────────────────┬────────────────────────────┘
                            │ Atomic disk persistence
                            ▼
┌────────────────────────────────────────────────────────┐
│  Output Layer (Atomic Staging & fsync)                 │
│  - Writes to hidden staging file: .{name}.{uuid}.tmp   │
│  - Explicit flush and physical disk sync via os.fsync  │
│  - Atomic replacement via os.replace                   │
└───────────────────────────┬────────────────────────────┘
                            │
        ┌───────────────────┼───────────────────┬───────────────────┐
        ▼                   ▼                   ▼                   ▼
┌───────────────┐   ┌───────────────┐   ┌───────────────┐   ┌───────────────┐
│  CLI (main)   │   │  FastAPI Web  │   │  Windows GUI  │   │ Report Engine │
│  deffilegen   │   │  web/app.py   │   │  gui.py       │   │ report_gen.py │
└───────────────┘   └───────────────┘   └───────────────┘   └───────────────┘
```

---

## ⚡ Key Architectural Invariants

1. **Lazy Row Pipeline**: Parsers expose iterators of tables and rows. Source bytes are buffered for document parsing; Excel archive sizes and expanded worksheet dimensions are bounded before expensive processing.
2. **Strict Determinism**: Generated output files are guaranteed to be byte-identical across runs with identical inputs, unaffected by `PYTHONHASHSEED`.
3. **Decoupled Responsibilities**: 
   - `extractor.py` handles document parsing and column mapping heuristics.
   - `knowledge_miner.py` handles domain semantic tagging, canonical mapping, and action classification.
   - `def_gen.py` handles data type semantics, address normalization, overlap checks, and CSV serialization.
   - `report_generator.py` handles multi-dimensional quality auditing, Markdown generation, and JSON metrics export.
   - Entry points (`main.py`, `gui.py`, `web/app.py`) depend on the core engine, never vice-versa.
4. **Defensive Parsing at System Boundaries**: Malformed XML entities, CSV formula injection attempts, and control characters are neutralized before reaching state.
5. **Conservative Completeness**: The pipeline never fabricates missing addresses. Scanned PDFs require external OCR.
6. **Transactional Publication**: Configured conversion validates complete staged output before publishing it. Input/validation failures emit no configured stdout CSV and preserve existing file destinations; caller-owned streams remain open.
7. **Network Storage Protection**: As mandated by project constitution, no tool or process may write, modify, or delete files on the shared network storage (`\\WDN-FILES.WEBDYN.COM`) without explicit confirmation from the user.

---

## 🌐 Public Surfaces & Entry Points

| Module / Script | Interface | Primary Purpose |
|---|---|---|
| [`DefFileGenerator/main.py`](../DefFileGenerator/main.py) | CLI (`deffilegen`) | Command-line interface (`run`, `extract`, `generate`, `validate`, `tag`, `audit`, `report`, `train`, `gui`) |
| [`DefFileGenerator/gui.py`](../DefFileGenerator/gui.py) | Desktop GUI (`deffilegen-gui`) | Native Windows 11 CustomTkinter application with live preview |
| [`web/app.py`](../web/app.py) | REST API / Web App | FastAPI backend and bundled modern HTML5/CSS3 frontend |
| [`DefFileGenerator/knowledge_miner.py`](../DefFileGenerator/knowledge_miner.py) | Semantic Domain Engine | SunPM canonical tagging, legacy normalization, and action code inference |
| [`DefFileGenerator/report_generator.py`](../DefFileGenerator/report_generator.py) | Audit & Report Generator | Technical KPI report formatting (Markdown & JSON) |
| [`DefFileGenerator/mars_generator.py`](../DefFileGenerator/mars_generator.py) | Industrial Generator | Mars Renewable BESS Local Controller V7.0 multi-slave definition generator |
| [`generate_webdyn_def.py`](../generate_webdyn_def.py) | Programmatic API | Primary Python API function: `generate_webdyn_definition(...)` |
| [`doc_to_webdyn.py`](../doc_to_webdyn.py) | Legacy CLI wrapper | Backward compatibility wrapper maintained for legacy scripts |
| [`scripts/bootstrap.ps1`](../scripts/bootstrap.ps1) | PowerShell Bootstrap | Automated environment setup and quality gate verification for Windows |
| [`scripts/bootstrap.sh`](../scripts/bootstrap.sh) | Bash Bootstrap | Automated environment setup and quality gate verification for Linux/macOS |

---

<p align="center">
  <a href="../README.md"><b>⬅ Back to README</b></a> •
  <a href="input-format.md"><b>Input Formats & Columns ➡</b></a>
</p>
