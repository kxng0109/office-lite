"""Build pptx bytes from blocks or csv rows (DRY, template-free)."""

from __future__ import annotations

import io
import threading

from pptx import Presentation
from pptx.util import Inches

from .guards import (
    MAX_CELLS,
    MAX_SLIDES,
    MAX_TABLE_COLS,
    MAX_TABLE_ROWS,
    MAX_TABLES,
    MAX_TOTAL_CELLS,
    check_size,
    sanitize_csv_cell,
)
from .md_parse import DocBlocks, TableBlock

MAX_BULLETS_PER_SLIDE = 20
MAX_CODE_CHARS = 2000
ROWS_PER_SLIDE = 15

_cached_template: bytes | None = None
_template_lock = threading.Lock()


def _new_deck() -> Presentation:
    """Fresh deck from a cached default template (bytes are re-parsed per call).

    The template bytes are cached to skip file I/O, but each call still builds
    a fresh in-memory package: python-pptx offers no safe way to clone a live
    presentation, so re-parsing is the price of isolation between requests.
    """
    global _cached_template
    if _cached_template is None:
        with _template_lock:
            if _cached_template is None:
                buf = io.BytesIO()
                Presentation().save(buf)
                _cached_template = buf.getvalue()
    prs = Presentation(io.BytesIO(_cached_template))
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    return prs


def build_pptx(blocks: DocBlocks, title: str = "Deck") -> bytes:
    prs = _new_deck()
    layout_title = prs.slide_layouts[0]
    layout_content = prs.slide_layouts[1]
    first = prs.slides.add_slide(layout_title)
    first.shapes.title.text = title
    chunks: list[tuple[str, list[str]]] = []
    current_title = title
    current_points: list[str] = []
    for level, text in blocks.headings:
        if level <= 2:
            if current_points or current_title != title:
                chunks.append((current_title, current_points))
            current_title = text
            current_points = []
        else:
            current_points.append(text)
    for para in blocks.paragraphs:
        current_points.append(para)
    for _, item in blocks.bullets:
        current_points.append(item)
    chunks.append((current_title, current_points))
    table_pages = sum(
        max((len(t.rows) + ROWS_PER_SLIDE - 1) // ROWS_PER_SLIDE, 1) for t in blocks.tables
    )
    if len(blocks.tables) > MAX_TABLES:
        raise ValueError("document exceeds table cap")
    if sum(len(t.header or []) * (len(t.rows) + 1) for t in blocks.tables) > MAX_TOTAL_CELLS:
        raise ValueError("document exceeds total cell cap")
    planned = len(chunks) + table_pages + len(blocks.codes) + 1
    if planned > MAX_SLIDES:
        raise ValueError("deck exceeds slide cap")
    if chunks and not chunks[-1][1] and chunks[-1][0] == title and len(chunks) > 1:
        chunks.pop()
    seen_content = False
    for slide_title, points in chunks:
        if not points and slide_title == title and seen_content:
            continue
        seen_content = True
        slide = prs.slides.add_slide(layout_content)
        slide.shapes.title.text = slide_title
        body = slide.placeholders[1].text_frame
        body.clear()
        for point in points[:MAX_BULLETS_PER_SLIDE]:
            para = body.add_paragraph()
            para.text = point
            para.level = 0
        if len(points) > MAX_BULLETS_PER_SLIDE:
            marker = body.add_paragraph()
            marker.text = "... (truncated)"
            marker.level = 0
    for table in blocks.tables:
        if len(table.header) * (len(table.rows) + 1) > MAX_CELLS:
            raise ValueError("table exceeds cell cap")
        for start in range(0, max(len(table.rows), 1), ROWS_PER_SLIDE):
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            _add_table(
                slide,
                TableBlock(header=table.header, rows=table.rows[start : start + ROWS_PER_SLIDE]),
            )
            if not table.rows:
                break
    for code in blocks.codes:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(6))
        text = code[:MAX_CODE_CHARS]
        if len(code) > MAX_CODE_CHARS:
            text += "\n... (truncated)"
        box.text_frame.text = text
    buf = io.BytesIO()
    prs.save(buf)
    data = buf.getvalue()
    check_size(data)
    return data


def build_pptx_from_rows(header: list[str], rows: list[list[str]], title: str = "Table") -> bytes:
    if len(header) > MAX_TABLE_COLS:
        raise ValueError("table exceeds column cap")
    if len(rows) > MAX_TABLE_ROWS:
        raise ValueError("table exceeds row cap")
    if len(header) * (len(rows) + 1) > MAX_CELLS:
        raise ValueError("table exceeds cell cap")
    prs = _new_deck()
    per_slide = ROWS_PER_SLIDE
    pages = max((len(rows) + per_slide - 1) // per_slide, 1)
    if pages > MAX_SLIDES:
        raise ValueError("deck exceeds slide cap")
    for start in range(0, max(len(rows), 1), per_slide):
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        if slide.shapes.title:
            slide.shapes.title.text = title
        window = rows[start : start + per_slide]
        _add_table(slide, TableBlock(header=header, rows=window))
        if not window:
            break
    buf = io.BytesIO()
    prs.save(buf)
    data = buf.getvalue()
    check_size(data)
    return data


def _add_table(slide, table: TableBlock) -> None:
    header = table.header or (
        [f"col{i + 1}" for i in range(len(table.rows[0]))] if table.rows else []
    )
    if not header:
        return
    if len(header) > MAX_TABLE_COLS:
        raise ValueError("table exceeds column cap")
    if len(table.rows) > MAX_TABLE_ROWS:
        raise ValueError("table exceeds row cap")
    if len(header) * (len(table.rows) + 1) > MAX_CELLS:
        raise ValueError("table exceeds cell cap")
    data_rows = table.rows
    shape = slide.shapes.add_table(
        len(data_rows) + 1, len(header), Inches(0.5), Inches(1.0), Inches(12), Inches(5)
    )
    tbl = shape.table
    for i, name in enumerate(header):
        tbl.cell(0, i).text = sanitize_csv_cell(name)
    for r, row in enumerate(data_rows, start=1):
        padded = (row + [""] * len(header))[: len(header)]
        for i, value in enumerate(padded):
            tbl.cell(r, i).text = sanitize_csv_cell(value)
