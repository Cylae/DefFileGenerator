# Export Security & Integrity

## CSV Injection Mitigation

A CSV cell opened in spreadsheet software (e.g. Microsoft Excel or LibreOffice Calc) can be executed as an active formula if its first significant character is a formula trigger. Spreadsheet software strips leading whitespace characters (`space`, `tab`, `CR`, `LF`, `NBSP`) *before* evaluating the formula, and normalizes fullwidth Unicode operators. A sanitization check restricted to raw ASCII first characters can therefore be bypassed.

The trigger set handled by `Generator.sanitize_csv_field`:

| Category | Characters |
|--|--|
| ASCII Triggers | `=` `+` `-` `@` `\|` `%` |
| Leading Whitespace | `0x20` `0x09` (tab) `0x0D` (CR) `0x0A` (LF) `0x0B` `0x0C` `U+00A0` (NBSP) |
| Fullwidth Variants | `＝` (`\uff1d`) `＋` (`\uff0b`) `－` (`\uff0d`) `＠` (`\uff20`) |

Any string value whose first non-whitespace character belongs to this trigger set has a single apostrophe (`'`) prepended in accordance with OWASP recommendations. Finite signed numbers (`-10.5`, `+25`, `1.5e3`) preserve their sign to retain numerical semantics; non-finite literals (`-inf`, `+nan`) and non-numeric string representations (` -10.5`, `-1_000`) are escaped.

Treat all vendor documentation as untrusted input data. Always validate generated definition CSV files prior to uploading them to field gateways or production devices.

## XML Entity & XXE Protection

XML parsers must remain protected against XML External Entity (XXE) attacks and billion-laughs expansion bombs. The extractor module strictly enforces `defusedxml.ElementTree` parsing (`DTDForbidden`, `EntitiesForbidden`, `ExternalReferenceForbidden`). Standard `xml.etree.ElementTree` without `defusedxml` protections is never used for untrusted files.

## Atomic File Writing

Generated definition files are written to a temporary staging file on the same filesystem volume before being atomically renamed into place. An interrupted or failing run will never corrupt or truncate an existing production definition file.

## Secret Management

No credentials, API tokens, customer exports, or sensitive site data must be stored in test fixtures. Synthetic dummy values are used throughout test suites.
