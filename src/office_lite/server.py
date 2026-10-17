"""Office-lite stdio MCP server: tools-only plus search."""

from __future__ import annotations

import base64
import binascii
import io
import logging
import sys
import zipfile

from docx import Document
from mcp.server.mcpserver import MCPServer
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from .csv_bridge import read_csv_text, rows_to_csv_text, rows_to_markdown
from .docx_build import build_docx, build_docx_from_rows
from .guards import (
    MAX_INPUT_BYTES,
    TRUNCATE_CHARS,
    audit_zip_bytes,
    check_outline,
    check_size,
    check_title,
    truncate_text,
)
from .md_parse import DocBlocks, parse_markdown, parse_text
from .pptx_build import build_pptx, build_pptx_from_rows

logger = logging.getLogger(__name__)

server: MCPServer = MCPServer(name="office-lite")

TOOLS_INDEX = [
    ("convert_markdown", "Convert markdown or plain text to docx or pptx bytes"),
    ("convert_table_in", "Convert csv text to docx, pptx, markdown table, or normalized csv"),
    ("convert_office_out", "Extract docx or pptx bytes to markdown, text, or csv"),
    ("create_from_outline", "Create docx or pptx from heading outline"),
    ("inspect_office", "Inspect docx or pptx bytes for counts without full convert"),
    ("search", "Search office-lite tools by query"),
]


def _decode_office_b64(file_b64: str) -> bytes:
    if len(file_b64) > MAX_INPUT_BYTES * 4 // 3 + 1024:
        raise ValueError("input exceeds size cap")
    try:
        raw = base64.b64decode(file_b64, validate=True)
    except (binascii.Error, ValueError) as exc:
        raise ValueError("invalid base64 input") from exc
    check_size(raw)
    audit_zip_bytes(raw)
    return raw


def _open_docx(raw: bytes):
    try:
        return Document(io.BytesIO(raw))
    except (zipfile.BadZipFile, KeyError, ValueError) as exc:
        raise ValueError("invalid docx file") from exc


def _open_pptx(raw: bytes):
    try:
        return Presentation(io.BytesIO(raw))
    except (zipfile.BadZipFile, KeyError, ValueError) as exc:
        raise ValueError("invalid pptx file") from exc


def _walk_shapes(shapes):
    """Yield shapes recursively, descending into group shapes."""
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _walk_shapes(shape.shapes)


@server.tool(
    title="Convert markdown", description="Convert markdown or plain text to docx or pptx bytes"
)
def convert_markdown(
    content: str, target: str, title: str = "Document", flavor: str = "markdown"
) -> str:
    if target not in ("docx", "pptx"):
        raise ValueError("target must be docx or pptx")
    if flavor not in ("markdown", "text"):
        raise ValueError("flavor must be markdown or text")
    check_title(title)
    blocks = parse_markdown(content) if flavor == "markdown" else parse_text(content)
    data = build_docx(blocks, title) if target == "docx" else build_pptx(blocks, title)
    return base64.b64encode(data).decode("ascii")


@server.tool(
    title="Convert table in", description="Convert csv text to docx, pptx, or markdown table"
)
def convert_table_in(csv_text: str, target: str, title: str = "Table", delimiter: str = ",") -> str:
    if target not in ("docx", "pptx", "markdown", "csv"):
        raise ValueError("target must be docx, pptx, markdown, or csv")
    check_title(title)
    header, rows = read_csv_text(csv_text, delimiter)
    if target == "markdown":
        return truncate_text(rows_to_markdown(header, rows, TRUNCATE_CHARS), TRUNCATE_CHARS)
    if target == "csv":
        return rows_to_csv_text(header, rows)
    data = (
        build_docx_from_rows(header, rows, title)
        if target == "docx"
        else build_pptx_from_rows(header, rows, title)
    )
    return base64.b64encode(data).decode("ascii")


@server.tool(
    title="Extract office", description="Extract docx or pptx bytes to markdown, text, or csv"
)
def convert_office_out(file_b64: str, source: str, target: str) -> str:
    if source not in ("docx", "pptx"):
        raise ValueError("source must be docx or pptx")
    if target not in ("markdown", "text", "csv"):
        raise ValueError("target must be markdown, text, or csv")
    raw = _decode_office_b64(file_b64)
    if source == "docx":
        doc = _open_docx(raw)
        paras = [p.text for p in doc.paragraphs if p.text.strip()]
        tables: list[str] = []
        for table in doc.tables:
            if not table.rows:
                continue
            header = [c.text for c in table.rows[0].cells]
            rows = [[c.text for c in r.cells] for r in table.rows[1:]]
            if target == "csv":
                tables.append(rows_to_csv_text(header, rows))
            else:
                tables.append(rows_to_markdown(header, rows, TRUNCATE_CHARS))
        if target == "csv":
            return truncate_text("\n".join(tables), TRUNCATE_CHARS)
        text = "\n\n".join(paras + tables)
        return truncate_text(text, TRUNCATE_CHARS)
    prs = _open_pptx(raw)
    slides: list[str] = []
    csv_parts: list[str] = []
    for i, slide in enumerate(prs.slides, 1):
        parts = [f"# Slide {i}"]
        for shape in _walk_shapes(slide.shapes):
            if shape.has_table:
                if not shape.table.rows:
                    continue
                header = [c.text for c in shape.table.rows[0].cells]
                rows = [[c.text for c in r.cells] for r in list(shape.table.rows)[1:]]
                if target == "csv":
                    csv_parts.append(rows_to_csv_text(header, rows))
                else:
                    parts.append(rows_to_markdown(header, rows, TRUNCATE_CHARS))
            elif shape.has_text_frame:
                if shape.text.strip():
                    parts.append(shape.text.strip())
        slides.append("\n\n".join(parts))
    if target == "csv":
        return truncate_text("\n".join(csv_parts), TRUNCATE_CHARS)
    return truncate_text("\n\n".join(slides), TRUNCATE_CHARS)


@server.tool(title="Create from outline", description="Create docx or pptx from heading outline")
def create_from_outline(headings: list[str], target: str, title: str = "Document") -> str:
    if target not in ("docx", "pptx"):
        raise ValueError("target must be docx or pptx")
    check_title(title)
    items = check_outline(headings)
    blocks = DocBlocks(headings=[(1, h) for h in items])
    data = build_docx(blocks, title) if target == "docx" else build_pptx(blocks, title)
    return base64.b64encode(data).decode("ascii")


@server.tool(
    title="Inspect office", description="Inspect docx or pptx bytes for counts without full convert"
)
def inspect_office(file_b64: str, source: str) -> str:
    if source not in ("docx", "pptx"):
        raise ValueError("source must be docx or pptx")
    raw = _decode_office_b64(file_b64)
    if source == "docx":
        doc = _open_docx(raw)
        return f"paragraphs={len(doc.paragraphs)} tables={len(doc.tables)} bytes={len(raw)}"
    prs = _open_pptx(raw)
    tables = sum(1 for s in prs.slides for sh in _walk_shapes(s.shapes) if sh.has_table)
    return f"slides={len(prs.slides)} tables={tables} bytes={len(raw)}"


@server.tool(title="Search tools", description="Search office-lite tools by query")
def search(query: str, limit: int = 10) -> str:
    if not isinstance(query, str):
        raise ValueError("query must be a string")
    q = query.strip().lower()[:200]
    if not q:
        raise ValueError("query is required")
    if isinstance(limit, bool) or not isinstance(limit, int):
        raise ValueError("limit must be an integer")
    limit = max(1, min(limit, 20))
    hits = [
        f"{name}: {desc}" for name, desc in TOOLS_INDEX if q in name.lower() or q in desc.lower()
    ]
    return "\n".join(hits[:limit]) if hits else "no matches"


def main() -> None:
    logging.basicConfig(
        level=logging.INFO, stream=sys.stderr, format="office-lite %(levelname)s %(message)s"
    )
    logger.info("starting office-lite stdio server")
    try:
        server.run()
    except Exception:
        logger.warning("server stopped with error")
        raise


if __name__ == "__main__":
    main()
