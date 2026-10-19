"""Office-lite stdio MCP server: tools-only plus search."""

from __future__ import annotations

import base64
import binascii
import io
import logging
import sys

from docx import Document
from mcp.server.mcpserver import MCPServer
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from .csv_bridge import read_csv_text, rows_to_csv_text, rows_to_markdown
from .docx_build import build_docx, build_docx_from_rows
from .guards import (
    MAX_GROUP_DEPTH,
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
    if not isinstance(file_b64, str):
        raise ValueError("file_b64 must be a string")
    if len(file_b64) > MAX_INPUT_BYTES * 4 // 3 + 1024:
        raise ValueError("input exceeds size cap")
    try:
        raw = base64.b64decode(file_b64, validate=True)
    except (binascii.Error, ValueError) as exc:
        raise ValueError("invalid base64 input") from exc
    check_size(raw)
    audit_zip_bytes(raw)
    logger.info("decode input_b64=%d bytes=%d", len(file_b64), len(raw))
    return raw


def _open_docx(raw: bytes):
    try:
        return Document(io.BytesIO(raw))
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError("invalid docx file") from exc


def _open_pptx(raw: bytes):
    try:
        return Presentation(io.BytesIO(raw))
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError("invalid pptx file") from exc


def _b64(data: bytes) -> str:
    # Output size is bounded transitively: base64 expands input by 4/3, and
    # check_size caps input, so a >13.4 MB response is unreachable by construction.
    check_size(data)
    text = base64.b64encode(data).decode("ascii")
    logger.info("encode bytes=%d b64=%d", len(data), len(text))
    return text


def _merge_tables(
    parts: list[tuple[list[str], list[list[str]]]],
) -> tuple[list[str], list[list[str]]]:
    """Merge same-header table fragments (e.g. paginated slides); reject mixed schemas."""
    if not parts:
        return [], []
    header = parts[0][0]
    if any(h != header for h, _ in parts):
        raise ValueError("multi-table csv export is unsupported")
    rows: list[list[str]] = []
    for _, block_rows in parts:
        rows.extend(block_rows)
    return header, rows


def _docx_para_markdown_cell(cell) -> str:
    """Cell text with hyperlinks preserved, paragraphs joined like cell.text."""
    return "\n".join(_docx_para_markdown(p) for p in cell.paragraphs)


def _walk_tables(tables, depth: int = 0):
    """Yield tables recursively, descending into nested cell tables with a depth cap."""
    if depth > MAX_GROUP_DEPTH:
        raise ValueError("table nesting exceeds depth cap")
    for table in tables:
        yield table
        for row in table.rows:
            for cell in row.cells:
                yield from _walk_tables(cell.tables, depth + 1)


def _w_t_text(elm) -> str:
    """Text from w:t descendants only; w:r direct text mirrors w:t in some writers."""
    texts = [t.text or "" for t in elm.iter() if isinstance(t.tag, str) and t.tag.endswith("}t")]
    if texts:
        return "".join(texts)
    return elm.text or ""


def _docx_para_markdown(paragraph) -> str:
    """Paragraph text with hyperlinks preserved as [text](url)."""
    try:
        links = {h._hyperlink: (h.text, h.url) for h in paragraph.hyperlinks}
    except Exception:
        links = {}
    parts: list[str] = []
    for child in paragraph._p.iterchildren():
        tag = child.tag.split("}")[-1] if isinstance(child.tag, str) else ""
        if tag == "hyperlink":
            text, url = links.get(child, (_w_t_text(child), ""))
            parts.append(f"[{text}]({url})" if url else text)
        elif tag == "r":
            parts.append(_w_t_text(child))
    text = "".join(parts)
    return text if text.strip() else paragraph.text


def _pptx_para_markdown(paragraph) -> str:
    """Text-frame paragraph with run hyperlinks preserved as [text](url)."""
    parts: list[str] = []
    try:
        runs = list(paragraph.runs)
    except Exception:
        return paragraph.text
    for run in runs:
        try:
            address = run.hyperlink.address
        except Exception:
            address = ""
        if address:
            parts.append(f"[{run.text}]({address})")
        else:
            parts.append(run.text)
    text = "".join(parts)
    return text if text.strip() else paragraph.text


def _walk_shapes(shapes, depth: int = 0):
    """Yield shapes recursively, descending into group shapes with a depth cap."""
    if depth > MAX_GROUP_DEPTH:
        raise ValueError("shape nesting exceeds depth cap")
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _walk_shapes(shape.shapes, depth + 1)


@server.tool(
    title="Convert markdown", description="Convert markdown or plain text to docx or pptx bytes"
)
def convert_markdown(
    content: str, target: str, title: str = "Document", flavor: str = "markdown"
) -> str:
    if not isinstance(content, str):
        raise ValueError("content must be a string")
    if target not in ("docx", "pptx"):
        raise ValueError("target must be docx or pptx")
    if flavor not in ("markdown", "text"):
        raise ValueError("flavor must be markdown or text")
    check_title(title)
    blocks = parse_markdown(content) if flavor == "markdown" else parse_text(content)
    data = build_docx(blocks, title) if target == "docx" else build_pptx(blocks, title)
    return _b64(data)


@server.tool(
    title="Convert table in", description="Convert csv text to docx, pptx, or markdown table"
)
def convert_table_in(csv_text: str, target: str, title: str = "Table", delimiter: str = ",") -> str:
    if not isinstance(csv_text, str):
        raise ValueError("csv_text must be a string")
    if target not in ("docx", "pptx", "markdown", "csv"):
        raise ValueError("target must be docx, pptx, markdown, or csv")
    check_title(title)
    header, rows = read_csv_text(csv_text, delimiter)
    if not rows and not any(cell.strip() for cell in header):
        raise ValueError("csv is empty")
    if target == "markdown":
        return truncate_text(rows_to_markdown(header, rows, TRUNCATE_CHARS), TRUNCATE_CHARS)
    if target == "csv":
        return rows_to_csv_text(header, rows)
    data = (
        build_docx_from_rows(header, rows, title)
        if target == "docx"
        else build_pptx_from_rows(header, rows, title)
    )
    return _b64(data)


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
        paras = [_docx_para_markdown(p) for p in doc.paragraphs]
        paras = [p for p in paras if p.strip()]
        tables: list[str] = []
        csv_blocks: list[tuple[list[str], list[list[str]]]] = []
        for table in _walk_tables(doc.tables):
            if not table.rows:
                continue
            header = [_docx_para_markdown_cell(c) for c in table.rows[0].cells]
            rows = [[_docx_para_markdown_cell(c) for c in r.cells] for r in table.rows[1:]]
            if target == "csv":
                csv_blocks.append((header, rows))
            else:
                tables.append(rows_to_markdown(header, rows, TRUNCATE_CHARS))
        if target == "csv":
            if not tables and not csv_blocks:
                return ""
            header, rows = _merge_tables(csv_blocks)
            return truncate_text(rows_to_csv_text(header, rows), TRUNCATE_CHARS)
        text = "\n\n".join(paras + tables)
        return truncate_text(text, TRUNCATE_CHARS)
    prs = _open_pptx(raw)
    slides: list[str] = []
    csv_blocks: list[tuple[list[str], list[list[str]]]] = []
    for i, slide in enumerate(prs.slides, 1):
        parts = [f"# Slide {i}"]
        for shape in _walk_shapes(slide.shapes):
            if shape.has_table:
                if not shape.table.rows:
                    continue
                header = [c.text for c in shape.table.rows[0].cells]
                rows = [[c.text for c in r.cells] for r in list(shape.table.rows)[1:]]
                if target == "csv":
                    csv_blocks.append((header, rows))
                else:
                    parts.append(rows_to_markdown(header, rows, TRUNCATE_CHARS))
            elif shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    rendered = _pptx_para_markdown(paragraph)
                    if rendered.strip():
                        parts.append(rendered.strip())
        slides.append("\n\n".join(parts))
    if target == "csv":
        if not csv_blocks:
            return ""
        header, rows = _merge_tables(csv_blocks)
        return truncate_text(rows_to_csv_text(header, rows), TRUNCATE_CHARS)
    return truncate_text("\n\n".join(slides), TRUNCATE_CHARS)


@server.tool(title="Create from outline", description="Create docx or pptx from heading outline")
def create_from_outline(headings: list[str], target: str, title: str = "Document") -> str:
    if target not in ("docx", "pptx"):
        raise ValueError("target must be docx or pptx")
    check_title(title)
    items = check_outline(headings)
    blocks = DocBlocks(headings=[(1, h) for h in items])
    data = build_docx(blocks, title) if target == "docx" else build_pptx(blocks, title)
    return _b64(data)


@server.tool(
    title="Inspect office", description="Inspect docx or pptx bytes for counts without full convert"
)
def inspect_office(file_b64: str, source: str) -> str:
    if source not in ("docx", "pptx"):
        raise ValueError("source must be docx or pptx")
    raw = _decode_office_b64(file_b64)
    if source == "docx":
        doc = _open_docx(raw)
        paras = len([p for p in doc.paragraphs if p.text.strip()])
        tables = sum(1 for _ in _walk_tables(doc.tables))
        return f"paragraphs={paras} tables={tables} bytes={len(raw)}"
    prs = _open_pptx(raw)
    tables = sum(1 for s in prs.slides for sh in _walk_shapes(s.shapes) if sh.has_table)
    return f"slides={len(prs.slides)} tables={tables} bytes={len(raw)}"


@server.tool(title="Search tools", description="Search office-lite tools by query")
def search(query: str, limit: int = 10) -> str:
    if not isinstance(query, str):
        raise ValueError("query must be a string")
    if len(query.strip()) > 200:
        raise ValueError("query exceeds length cap")
    q = query.strip().lower()
    if not q:
        raise ValueError("query is required")
    if isinstance(limit, bool) or not isinstance(limit, int):
        raise ValueError("limit must be an integer")
    if limit < 1 or limit > 20:
        raise ValueError("limit must be between 1 and 20")
    hits = [
        f"{name}: {desc}" for name, desc in TOOLS_INDEX if q in name.lower() or q in desc.lower()
    ]
    return "\n".join(hits[:limit]) if hits else "no matches"


def main() -> None:
    logging.basicConfig(
        level=logging.INFO,
        stream=sys.stderr,
        format="%(asctime)s office-lite %(levelname)s %(message)s",
    )
    logger.info("starting office-lite stdio server")
    try:
        server.run()
    except Exception:
        logger.exception("server stopped with error")
        raise SystemExit(1) from None


if __name__ == "__main__":
    main()
