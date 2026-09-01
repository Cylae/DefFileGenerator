# Changelog

All notable changes to this project are documented in this file.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and the project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.1]

Rebuilt against the current `main` (`0167102`). Functionally identical to the
0.2.0 audit; the version was incremented because that release was never merged
upstream and this archive supersedes it.

### Security

- **Formula-injection sanitiser hardened.** `sanitize_csv_field` previously
  inspected only the raw first byte. Spreadsheet clients strip leading
  whitespace before deciding whether a cell is a formula and normalise
  full-width punctuation to ASCII, so payloads prefixed with a tab, CR, LF,
  vertical tab, form feed or NBSP — and full-width `＝ ＋ － ＠` — were written
  unescaped. The DDE pipe (`|`) was not treated as a trigger at all. Escaping
  is now decided on the first *significant* character. The numeric bypass
  additionally requires `math.isfinite()` and a strict decimal pattern, so
  `-inf`, `+nan` and `-1_000` are escaped while `-10.5`, `+25` and `1.5e3`
  remain verbatim. This closes a path from untrusted vendor documentation to
  code execution on an engineer's workstation.

### Fixed

- **Non-reproducible output.** Column detection iterated a `set` of header
  names. Because string hashing is randomised per interpreter, targets sharing
  a pattern (`offset` matches both Address and Offset) could bind to different
  source columns between runs — the same input produced different definition
  files, and rows were silently dropped on some seeds. Header order is now
  preserved.
- **Action code `10` was discarded.** `Constant` variables (firmware 5.2.02+)
  were rewritten to `1`, changing device behaviour. The code is now preserved,
  with the fallback for genuinely unknown codes unchanged.
- **Non-atomic writes.** A failure mid-stream truncated an existing definition
  file. Output is now staged via `tempfile.mkstemp` on the same filesystem and
  moved into place with `os.replace`; no temporary file survives a failure.

### Added

- Post-generation validation for `run`, with `--no-validate` to opt out.
- `--force` guard: an existing output file is no longer silently overwritten.
- `--lenient` for `validate`, downgrading address overlaps to warnings.
- `-q/--quiet` and `--version`; `-v`/`-q` are accepted before *or* after the
  sub-command.
- Actionable diagnostics: unsupported extensions list the supported formats,
  missing flags are named with a worked example, and empty extractions suggest
  `--sheet`, `--pages` or `--mapping`.
- Help output gained per-command descriptions, metavars, worked examples and
  documented exit codes (`0` success, `1` error, `2` usage, `130` interrupt).
- 26 regression tests covering injection payloads, atomic-write recovery,
  action-code parity, type/address precedence, hash-seed determinism and CLI
  ergonomics (77 → 103).

### Changed

- Register processing throughput improved ~1.46x on a 50,000-register map.
  `normalize_type` is substantially faster: twelve sequential `re.search` calls
  per row were replaced by a precompiled, specificity-ordered table behind an
  `lru_cache`. Address normalisation and type-width lookup are likewise
  memoised, action codes moved to `frozenset`, and `bisect` was hoisted out of
  the hot loop.
- Column detection reduced from O(K x T x P) to O(K) by lowercasing each header
  once instead of per probe.
- Behaviour is unchanged: normalisation parity was verified against the previous
  implementation over 600,000 random tokens with zero divergence, and all
  fixture outputs are byte-identical.
