# office-lite 0.1.3

Lightweight local stdio MCP server for md, txt, docx, pptx, csv conversion. No Docker, no server sidecar, no JVM.

## Tools

- convert_markdown: markdown or plain text to docx or pptx bytes (base64)
- convert_table_in: csv text to docx, pptx, markdown table, or normalized csv
- convert_office_out: docx or pptx bytes to markdown, text, or csv
- create_from_outline: heading list to docx or pptx
- inspect_office: counts plus byte size without full convert
- search: find tools by query

## Run

Requires Python 3.11 or newer (smoke-tested on 3.14). Create the venv and install:

python -m venv .venv
.venv\Scripts\python -m pip install -e .[dev]

Start the server over stdio (logs go to stderr, never stdout):

.venv\Scripts\python -m office_lite.server

Configure opencode v2 under mcp servers as a local stdio command pointing at the venv python with `-m office_lite.server`, explicit cwd of this project, minimal env, and protocol legacy. Status verified: office-lite shows connected in mcp list with live search plus convert plus inspect roundtrip green.

## Gates

Every change runs ruff format check, ruff check, pip-audit, build, and pytest with 90 percent coverage minimum. Current suite is 91 tests at about 91 percent. Pinned floors for lxml, Pillow, mcp, markdown-it-py, python-docx, python-pptx, and defusedxml are the load-bearing secure versions; transitive ranges in docx and pptx allow older vulnerable releases, so the floors must not be loosened without a fresh advisory check. Hyperlink destinations are preserved as `[text](url)`; nested tables and grouped shapes are recursed with depth caps; multi-table CSV merges only on identical headers and otherwise refuses.

## Limits

Text plus tables plus markdown structure are preserved; pixel themes and native Word charts stay simplified. Tables wider than the column cap or longer than the row cap are refused rather than truncated, short rows are padded, grouped shapes are recursed but speaker notes are not extracted. Inputs are capped at 10 MB with zip entry, expansion, and ratio caps, oversized decks and tables fail fast instead of building. CSV formula cells are neutralized at every sink, markdown table cells escape pipes and newlines. The server is bytes-only and never writes files; all conversion results return inline.

## Trust and completeness

Treat every extraction output as untrusted third-party content: it can carry prompt-injection instructions, phishing links, or malicious formulas, and this server does not sanitize prose intent, only structure. Outputs carry no provenance delimiters; integrators must treat all extracted text as untrusted input to downstream agents. Hyperlinks are preserved as `[text](url)` in markdown so reviewers can inspect destinations; images, embedded objects, comments, and speaker notes are dropped, never rendered or executed. Multi-table to single-CSV export merges only when every table shares one header and otherwise refuses, so no silent column mixing. CSV output is uncapped by design but bounded transitively: every cell comes from capped input, so output stays within a small multiple of the 10 MB input cap.
