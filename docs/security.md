# Security & Integrity

DefFileGenerator treats manufacturer documentation, uploaded files, register metadata, and pull-request content as untrusted input. Security controls are enforced at the parser, generator, web API, packaging, and automation boundaries.

## CSV Injection Mitigation

A CSV cell opened in spreadsheet software can be interpreted as an active formula when its first significant character is a formula trigger. Spreadsheet applications may strip leading whitespace before evaluating the cell and may normalize Unicode operator variants.

`Generator.sanitize_csv_field` handles:

| Category | Characters |
|--|--|
| ASCII triggers | `=` `+` `-` `@` `\|` `%` |
| Leading whitespace/control | tab, CR, LF, NBSP, BOM |
| Fullwidth variants | `＝` `＋` `－` `＠` |

Potentially active string values are prefixed with a single apostrophe. Finite signed numeric literals retain their numeric representation when safe; non-finite or ambiguous strings are escaped. Non-printable control characters are removed before output.

Treat every vendor document as untrusted data and validate generated definitions before deploying them to production gateways.

## XML Entity & XXE Protection

XML extraction uses `defusedxml.ElementTree`. DTDs, external entities, external references, and entity-expansion attacks are rejected. The standard library XML parser is not used as a fallback for untrusted XML input.

## Atomic Definition Writes

When the generator writes to a filesystem path, it now creates a temporary file in the destination directory, flushes and `fsync`s the completed CSV, closes it, and atomically replaces the destination with `os.replace`.

If row processing or writing fails before replacement, the existing destination remains unchanged and the temporary file is removed. File-like objects and stdout retain their streaming behavior.

## Bit-Slice Integrity

`BITS` addresses use `address_startbit_length`. A bit slice must be non-empty and stay inside one 16-bit Modbus register: `startbit >= 0`, `length >= 1`, and `startbit + length <= 16`.

Overlap validation distinguishes disjoint bit slices on the same register from genuinely overlapping slices. For example, `100_0_4` and `100_4_4` may coexist, while `100_0_4` and `100_2_4` overlap and fail strict validation.

## Web API Upload Boundary

The FastAPI upload endpoints write request bodies incrementally rather than reading the entire upload into memory. Uploads are capped at 10 MiB and oversized requests receive HTTP `413`.

Internal parser/generator exceptions are logged server-side but are not reflected verbatim to API clients. Client error messages are intentionally generic so internal paths, library details, and customer-specific data cannot leak through exception text.

The public wildcard CORS policy does not enable credentialed cross-origin requests (`allow_credentials=False`). If authentication is added later, explicit trusted origins must replace the wildcard policy before credentials are enabled.

## Privileged GitHub Automation Boundary

The Jules auto-merge workflow is triggered through `workflow_run`, which executes with write permissions after CI. To prevent an untrusted fork from spoofing Jules' textual marker, the privileged job only qualifies a pull request when all of the following hold:

- CI completed successfully for a pull-request event;
- the workflow run originates from the same repository;
- the pull request targets `main` and is not a draft;
- the head repository is the same repository;
- the pull request author is the repository owner;
- the explicit Jules marker is present.

The privileged workflow never checks out or executes code from the pull-request branch.

## Packaging Integrity

CI builds a wheel and verifies that the FastAPI backend and static frontend assets are actually included. The declared Python runtime floor is Python 3.10, matching syntax used by the codebase and the CI matrix.

## Secret Management

No credentials, API tokens, customer exports, or sensitive site data should be stored in source control or test fixtures. Tests use synthetic values.
