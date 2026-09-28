# Security & Integrity Model

<p align="center">
  <a href="../README.md"><b>README</b></a> •
  <a href="architecture.md"><b>Architecture</b></a> •
  <a href="development.md"><b>Development</b></a> •
  <a href="../AUDIT_REPORT.md"><b>Audit Report</b></a> •
  <a href="../DEVELOPER_GUIDE.md"><b>Developer Guide</b></a>
</p>

---

DefFileGenerator treats all manufacturer documentation files, uploaded spreadsheets/PDFs, user-provided metadata, and external PR artifacts as **untrusted input**. Rigorous security controls and sanitization policies are enforced at every architectural boundary.

---

## 📑 Table of Contents

- [🛡️ Security Controls Matrix](#️-security-controls-matrix)
- [💉 CSV Formula Injection Mitigation (DDE Defense)](#-csv-formula-injection-mitigation-dde-defense)
- [📦 XML Entity & XXE Attack Protection](#-xml-entity--xxe-attack-protection)
- [💾 Atomic Output Writes & Filesystem Resilience](#-atomic-output-writes--filesystem-resilience)
- [🧩 Bitfield Slice Integrity (`BITS`)](#-bitfield-slice-integrity-bits)
- [🌐 Web API Ingestion Boundaries (FastAPI)](#-web-api-ingestion-boundaries-fastapi)
- [🤖 Hardened Privileged GitHub Actions Automation](#-hardened-privileged-github-actions-automation)

---

## 🛡️ Security Controls Matrix

| Attack Vector / Risk | Defensive Mechanism | Enforcing Component |
|---|---|---|
| **CSV Formula Injection** | Apostrophe `'` prefix on trigger characters + Unicode normalization | `Generator.sanitize_csv_field` |
| **XML External Entity (XXE)** | DTD and external entity parsing disabled via `defusedxml` | `Extractor._extract_xml` |
| **Partial Output Corruption** | Staged sibling tempfile + `os.fsync` + atomic `os.replace` | `Generator.generate` |
| **Memory Exhaustion Denial of Service** | Strict 10 MiB upload cap + chunk streaming + 65,536 register limit | `web/app.py` |
| **Information Disclosure (CWE-209)** | Sanitized client error messages masking internal exception traces | `web/app.py` |
| **CI Automation Spoofing** | Head repo, author, branch, and cryptographic marker checks in Jules | `.github/workflows/` |

---

## 💉 CSV Formula Injection Mitigation (DDE Defense)

When a CSV file is opened in spreadsheet software (Microsoft Excel, LibreOffice Calc), any cell starting with an active formula trigger can execute dynamic spreadsheet macros or trigger system commands via Dynamic Data Exchange (DDE).

The `Generator.sanitize_csv_field` method inspects every cell string before writing:

| Category | Neutralized Trigger Characters |
|---|---|
| **Standard ASCII Triggers** | `=` `+` `-` `@` `\|` `%` |
| **Leading Whitespace & Hidden Characters** | Tab (`\t`), Carriage Return (`\r`), Line Feed (`\n`), Non-breaking space (`\u00A0`), UTF-8 BOM |
| **Fullwidth Unicode Variants** | `＝` (`\uFF1D`), `＋` (`\uFF0B`), `－` (`\uFF0D`), `＠` (`\uFF20`) |

> [!NOTE]
> **Preservation of Numeric Literals:**  
> Genuine signed numeric literals (e.g. `-10`, `+5.4`) are recognized and preserved as raw numeric values without an unnecessary leading apostrophe.

---

## 📦 XML Entity & XXE Attack Protection

XML files are parsed exclusively using `defusedxml.ElementTree`:
- External DTD declarations and entity expansions are systematically rejected.
- Recursive entity expansion attacks ("Billion Laughs") trigger an immediate defensive exception without consuming CPU or memory resources.

---

## 💾 Atomic Output Writes & Filesystem Resilience

When generating an output file on disk:
1. Data is written to a unique hidden temporary file in the destination directory: `.{name}.{uuid}.tmp`.
2. All buffers are flushed and physically synced to persistent storage using `os.fsync(fd)`.
3. The temporary file atomically replaces the destination path via `os.replace()`.

> [!IMPORTANT]
> If extraction fails, validation reports errors, or the process is abruptly terminated, the pre-existing target file remains **completely untouched and uncorrupted**, and the temporary file is safely cleaned up.

---

## 🧩 Bitfield Slice Integrity (`BITS`)

A `BITS` slice (`address_startbit_length`) must strictly fit within a single 16-bit Modbus word:
- `startbit` $\ge 0$
- `length` $\ge 1$
- `startbit + length` $\le 16$

The $O(\log N)$ interval algorithm verifies that multiple bitfield slices on the same base register address occupy disjoint bit ranges.

---

## 🌐 Web API Ingestion Boundaries (FastAPI)

The FastAPI web backend enforces defensive boundaries:
- **Maximum File Size**: Capped at **10 MiB**. Oversized payloads are rejected immediately with HTTP `413 Request Entity Too Large`.
- **Chunk-Based Streaming**: Uploads are processed in 64 KiB chunks, preventing large files from exhausting server RAM.
- **Register Ceiling**: Conversions reject files containing more than 65,536 registers.
- **Sanitized Client Errors**: Internal stack traces and server file paths are logged server-side and never returned to clients.

---

## 🤖 Hardened Privileged GitHub Actions Automation

The automated Jules merge workflow triggers via `workflow_run`. To prevent untrusted forks from exploiting write permissions, the job verifies that:
1. CI tests completed with 100% success.
2. The workflow originated from the root repository (`head_repository == repository`).
3. The pull request targets `main` and is not a draft.
4. The pull request author is the verified repository owner.
5. The explicit cryptographic Jules marker is present in metadata.

---

<p align="center">
  <a href="../README.md"><b>⬅ Back to README</b></a> •
  <a href="development.md"><b>Development Guidelines ➡</b></a>
</p>
