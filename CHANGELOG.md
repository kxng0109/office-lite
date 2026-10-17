# Changelog

All notable changes follow Semantic Versioning.

## [0.1.0] - 2026-10-17

### Added

- Initial office-lite stdio MCP server with six tools plus search.
- DRY shared guards for size caps, truncation, path jail, and CSV sanitize.
- Markdown and text parsing, csv bridge with streaming and BOM handling, docx and pptx builders, office extraction and inspection.
- Pinned light dependencies with smoke version log on Python 3.14.
- Test suite with adversarial cases covering empty, oversized, malformed, injection, traversal, encoding, merged tables, pagination, and outline limits.
- Lint, audit, build, and 90 percent coverage gates green.
