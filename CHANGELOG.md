# Changelog

All notable changes follow Semantic Versioning.

## [0.1.4] - 2026-10-17

### Fixed

- Unicode strip set now covers U+2028/29, NEL, VT/FF, Mongolian, word joiner, soft hyphen, and fullwidth no-break space, with per-codepoint tests.
- Hyperlink destinations preserved on docx and pptx extraction; nested tables and group shapes recursed with depth caps.
- Fail-closed wide-table, total-cell, and line-length caps on every ingestion and build path; csv delimiters allowlisted.
- Per-request stderr logging with sizes and durations; template bytes cached across builds.

## [0.1.3] - 2026-10-17

### Fixed

- Sink-side formula neutralization for docx and pptx builds; hyperlink destinations preserved on extraction.
- Generated Unicode strip set covering zero-width and format smugglers, with exhaustive per-codepoint tests.
- Rendering recursion depth caps, table-count and total-cell budgets, fail-closed row caps on every ingestion path.
- Strict input typing on every tool, control-character titles rejected but emoji-safe, search contract strict.
- Removed dead path-jail helpers; server is bytes-only and never writes files.
- Single-source version, corrected changelog math wording.
- Wide tables fail closed on every path; csv field sizes bounded by the stdlib field limit.
- Single-pass size checks, cached default templates, per-request stderr logging.
- Omitted-row counts corrected; silent truncation paths documented or removed.

## [0.1.2] - 2026-10-17

### Fixed

- Fail-closed table caps on both builders with pagination for block tables.
- Unicode-whitespace and zero-width formula smuggling neutralized.
- Hyperlink destinations preserved as `[text](url)` in markdown output.
- Multi-table CSV merges only on identical headers, otherwise refuses.
- Strict search contract, group depth cap, per-item length caps, title control-character reject.
- Dependency ceilings guard the next breaking major.
- Trust and completeness section documents untrusted-content handling.

## [0.1.1] - 2026-10-17

### Fixed

- Zip audit normalizes corrupt archives to ValueError, scans entries with an 8 MB audit-read budget, and uses integer ratio math with no float precision loss.
- Markdown tables sanitize formula cells and escape structure, csv delimiter allowlist is strict, and empty csv is rejected.
- Block tables paginate per slide in pptx and enforce a pre-build cell cap in docx.
- Contracts hardened: strict flavors, sources, titles, outlines, search limits, single CSV convention, and grouped-shape depth cap.
- Logging goes to stderr with timestamps and startup failures exit non-zero.
- Added LICENSE (MIT) and requirements.lock.

## [0.1.0] - 2026-10-17

### Added

- Initial office-lite stdio MCP server with six tools plus search.
- DRY shared guards for size caps, truncation, path jail, and CSV sanitize.
- Markdown and text parsing, csv bridge with streaming and BOM handling, docx and pptx builders, office extraction and inspection.
- Pinned light dependencies with smoke import check on Python 3.14.
- Test suite with adversarial cases covering empty, oversized, malformed, injection, traversal, encoding, merged tables, pagination, and outline limits.
- Lint, audit, build, and 90 percent coverage gates green.
