# Changelog

All notable changes follow Semantic Versioning.

## [0.1.1] - 2026-10-17

### Fixed

- Zip audit normalizes corrupt archives to ValueError, scans entries with an 8 MB audit-read budget, and uses float ratio math.
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
