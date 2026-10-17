# office-lite 0.1.0

Lightweight local stdio MCP server for md, txt, docx, pptx, csv conversion. No Docker, no server sidecar, no JVM.

## Tools

- convert_markdown: markdown or plain text to docx or pptx bytes (base64)
- convert_table_in: csv text to docx, pptx, markdown table, or normalized csv
- convert_office_out: docx or pptx bytes to markdown, text, or csv
- create_from_outline: heading list to docx or pptx
- inspect_office: counts plus byte size without full convert
- search: find tools by query

## Run

Use the project virtual environment on Windows with the MCP Python SDK v2. The server runs over stdio with logs on stderr. Configure opencode v2 under mcp servers as a local stdio command pointing at the venv python and the server module, with explicit cwd, minimal env, and protocol legacy. Status verified: office-lite shows connected in mcp list with live search plus convert plus inspect roundtrip green.

## Gates

Every change runs lint, dependency audit, build, and tests with 90 percent coverage minimum. Current suite is 21 tests at about 96 percent.

## Limits

Text plus tables plus markdown structure are preserved; pixel themes and native Word charts stay simplified. Extra columns beyond the header are dropped and short rows are padded, grouped shapes are recursed but speaker notes are not extracted. Inputs are capped at 10 MB with zip entry, expansion, and ratio caps, oversized decks and tables fail fast instead of building. CSV formula cells are neutralized, markdown table cells escape pipes and newlines. The server is bytes-only and never writes files; all conversion results return inline.
