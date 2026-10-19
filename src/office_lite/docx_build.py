"""Build docx bytes from blocks or csv rows (DRY, no template required)."""

from __future__ import annotations

import io
import threading

from docx import Document

from .guards import MAX_CELLS, MAX_TABLE_COLS, MAX_TABLES, MAX_TOTAL_CELLS, check_size
from .md_parse import DocBlocks, TableBlock

_cached_template: bytes | None = None
_template_lock = threading.Lock()


def _new_doc() -> Document:
    """Fresh document from a cached default template (bytes re-parsed per call)."""
    global _cached_template
    if _cached_template is None:
        with _template_lock:
            if _cached_template is None:
                buf = io.BytesIO()
                Document().save(buf)
                _cached_template = buf.getvalue()
    return Document(io.BytesIO(_cached_template))


def build_docx(blocks: DocBlocks, title: str = "Document") -> bytes:
    doc = _new_doc()
    if title:
        doc.add_heading(title, 0)
    for level, text in blocks.headings:
        doc.add_heading(text, min(max(level, 1), 9))
    for para in blocks.paragraphs:
        doc.add_paragraph(para)
    for _, item in blocks.bullets:
        doc.add_paragraph(item, style="List Bullet")
    for code in blocks.codes:
        doc.add_paragraph(code, style="Normal")
    if len(blocks.tables) > MAX_TABLES:
        raise ValueError("document exceeds table cap")
    if sum(len(t.header or []) * (len(t.rows) + 1) for t in blocks.tables) > MAX_TOTAL_CELLS:
        raise ValueError("document exceeds total cell cap")
    for table in blocks.tables:
        _add_table(doc, table)
    buf = io.BytesIO()
    doc.save(buf)
    data = buf.getvalue()
    check_size(data)
    return data


def build_docx_from_rows(header: list[str], rows: list[list[str]], title: str = "Table") -> bytes:
    doc = _new_doc()
    if title:
        doc.add_heading(title, 1)
    _add_table(doc, TableBlock(header=header, rows=rows))
    buf = io.BytesIO()
    doc.save(buf)
    data = buf.getvalue()
    check_size(data)
    return data


def _add_table(doc: Document, table: TableBlock) -> None:
    if not table.header and not table.rows:
        return
    header = (
        table.header or [f"col{i + 1}" for i in range(len(table.rows[0]))] if table.rows else []
    )
    if not header:
        return
    if len(header) > MAX_TABLE_COLS:
        raise ValueError("table exceeds column cap")
    if len(header) * (len(table.rows) + 1) > MAX_CELLS:
        raise ValueError("table exceeds cell cap")
    grid = doc.add_table(rows=1, cols=len(header))
    grid.style = "Light Grid Accent 1"
    for i, name in enumerate(header):
        grid.cell(0, i).text = name
    for row in table.rows:
        padded = (row + [""] * len(header))[: len(header)]
        cells = grid.add_row().cells
        for i, value in enumerate(padded):
            cells[i].text = value
