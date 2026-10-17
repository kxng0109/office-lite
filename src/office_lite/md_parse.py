"""Markdown and text parsing to block stream (DRY for docx/pptx/md)."""

from __future__ import annotations

import threading
from dataclasses import dataclass, field

from markdown_it import MarkdownIt

from .guards import (
    MAX_CODE_BLOCKS,
    MAX_HEADING_CHARS,
    MAX_HEADINGS,
    MAX_INPUT_BYTES,
    MAX_LINE_CHARS,
    MAX_LINES,
    MAX_LIST_DEPTH,
    MAX_PARA_CHARS,
    MAX_TABLE_COLS,
    MAX_TABLE_ROWS,
    MAX_TABLES,
    MAX_TEXT_BLOCKS,
    MAX_TOTAL_CELLS,
    check_size,
)

_local = threading.local()
_init_lock = threading.RLock()


def get_parser() -> MarkdownIt:
    """Thread-local parser; stdio transport serves one request at a time."""
    parser = getattr(_local, "parser", None)
    if parser is None:
        with _init_lock:
            parser = getattr(_local, "parser", None)
            if parser is None:
                parser = MarkdownIt("commonmark").enable("table")
                _local.parser = parser
    return parser


@dataclass
class TableBlock:
    header: list[str]
    rows: list[list[str]]


@dataclass
class DocBlocks:
    headings: list[tuple[int, str]] = field(default_factory=list)
    paragraphs: list[str] = field(default_factory=list)
    bullets: list[tuple[int, str]] = field(default_factory=list)
    codes: list[str] = field(default_factory=list)
    tables: list[TableBlock] = field(default_factory=list)


def _inline_text(token) -> str:
    if token.children:
        parts: list[str] = []
        i = 0
        children = token.children
        while i < len(children):
            child = children[i]
            if child.type == "link_open":
                href = (child.attrs or {}).get("href", "") if child.attrs else ""
                i += 1
                text_parts: list[str] = []
                while i < len(children) and children[i].type != "link_close":
                    if children[i].content:
                        text_parts.append(children[i].content)
                    i += 1
                label = "".join(text_parts)
                parts.append(f"[{label}]({href})" if href else label)
                i += 1
                continue
            if child.content:
                parts.append(child.content)
            i += 1
        return "".join(parts)
    return token.content or ""


def _check_text_size(text: str) -> None:
    if text.count("\n") + 1 > MAX_LINES:
        raise ValueError("input exceeds line cap")
    for line in text.splitlines():
        if len(line) > MAX_LINE_CHARS:
            raise ValueError("input line exceeds length cap")
    try:
        check_size(text.encode("utf-8"))
    except UnicodeEncodeError as exc:
        raise ValueError("invalid text encoding") from exc


def _consume_list(tokens, i: int, stack: list[list], blocks: DocBlocks) -> int:
    """Consume one markdown list including nested lists; preserves document order."""
    if len(stack) >= MAX_LIST_DEPTH:
        raise ValueError("list nesting exceeds depth cap")
    ordered = tokens[i].type == "ordered_list_open"
    stack.append(["ordered" if ordered else "bullet", 0])
    depth = len(stack) - 1
    i += 1
    while i < len(tokens) and tokens[i].type not in ("bullet_list_close", "ordered_list_close"):
        if tokens[i].type != "list_item_open":
            i += 1
            continue
        if ordered:
            stack[-1][1] += 1
        prefix = f"{stack[-1][1]}. " if ordered else ""
        slot = len(blocks.bullets)
        blocks.bullets.append((depth, ""))
        i += 1
        parts: list[str] = []
        while i < len(tokens) and tokens[i].type != "list_item_close":
            if tokens[i].type == "paragraph_open":
                inline = tokens[i + 1] if i + 1 < len(tokens) else None
                if inline is not None:
                    parts.append(_inline_text(inline))
                i += 3
            elif tokens[i].type in ("bullet_list_open", "ordered_list_open"):
                i = _consume_list(tokens, i, stack, blocks)
            else:
                i += 1
        text = " ".join(p for p in parts if p)
        blocks.bullets[slot] = (depth, (prefix + text) if text else "")
        i += 1
    stack.pop()
    return i + 1


def parse_markdown(text: str) -> DocBlocks:
    _check_text_size(text)
    blocks = DocBlocks()
    tokens = get_parser().parse(text)
    i = 0
    list_stack: list[list] = []
    while i < len(tokens):
        tok = tokens[i]
        if tok.type == "heading_open":
            level = int(tok.tag[1]) if tok.tag and len(tok.tag) == 2 else 1
            inline = tokens[i + 1] if i + 1 < len(tokens) else None
            title = _inline_text(inline) if inline is not None else ""
            blocks.headings.append((level, title))
            i += 3
        elif tok.type == "paragraph_open":
            inline = tokens[i + 1] if i + 1 < len(tokens) else None
            if inline is not None:
                blocks.paragraphs.append(_inline_text(inline))
            i += 3
        elif tok.type in ("bullet_list_open", "ordered_list_open"):
            i = _consume_list(tokens, i, list_stack, blocks)
        elif tok.type in ("fence", "code_block"):
            blocks.codes.append(tok.content or "")
            i += 1
        elif tok.type == "table_open":
            header: list[str] = []
            rows: list[list[str]] = []
            j = i + 1
            in_head = False
            current: list[str] = []
            while j < len(tokens) and tokens[j].type != "table_close":
                t = tokens[j]
                if t.type == "thead_open":
                    in_head = True
                elif t.type == "thead_close":
                    in_head = False
                elif t.type == "th_open":
                    inline = tokens[j + 1] if j + 1 < len(tokens) else None
                    cell = _inline_text(inline) if inline is not None else ""
                    current.append(cell)
                elif t.type == "td_open":
                    inline = tokens[j + 1] if j + 1 < len(tokens) else None
                    cell = _inline_text(inline) if inline is not None else ""
                    current.append(cell)
                elif t.type == "tr_close":
                    if in_head:
                        header = current[:MAX_TABLE_COLS]
                    else:
                        rows.append(current[:MAX_TABLE_COLS])
                        if len(rows) > MAX_TABLE_ROWS:
                            raise ValueError("table exceeds row cap")
                    current = []
                j += 1
            if len(blocks.tables) >= MAX_TABLES:
                raise ValueError("document exceeds table cap")
            blocks.tables.append(TableBlock(header=header, rows=rows))
            i = j + 1
        else:
            i += 1
    _enforce_caps(blocks)
    return blocks


def _enforce_caps(blocks: DocBlocks) -> None:
    if len(blocks.headings) > MAX_HEADINGS:
        raise ValueError("document exceeds heading cap")
    if any(len(text) > MAX_HEADING_CHARS for _, text in blocks.headings):
        raise ValueError("heading exceeds length cap")
    if len(blocks.paragraphs) + len(blocks.bullets) > MAX_TEXT_BLOCKS:
        raise ValueError("document exceeds text-block cap")
    if any(len(text) > MAX_PARA_CHARS for text in blocks.paragraphs):
        raise ValueError("paragraph exceeds length cap")
    if any(len(item) > MAX_PARA_CHARS for _, item in blocks.bullets):
        raise ValueError("list item exceeds length cap")
    if len(blocks.codes) > MAX_CODE_BLOCKS:
        raise ValueError("document exceeds code-block cap")
    if any(len(code) > MAX_INPUT_BYTES for code in blocks.codes):
        raise ValueError("code block exceeds length cap")
    if len(blocks.tables) > MAX_TABLES:
        raise ValueError("document exceeds table cap")
    total_cells = sum(len(t.header) * (len(t.rows) + 1) for t in blocks.tables)
    if total_cells > MAX_TOTAL_CELLS:
        raise ValueError("document exceeds total cell cap")


def parse_text(text: str) -> DocBlocks:
    _check_text_size(text)
    blocks = DocBlocks()
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        if line.startswith("#"):
            stripped = line.lstrip("#")
            if stripped.startswith((" ", "\t")) and stripped.strip():
                level = len(line) - len(line.lstrip("#"))
                blocks.headings.append((min(level, 6), stripped.strip()))
            else:
                blocks.paragraphs.append(line)
        elif line.startswith(("- ", "* ", "+ ")):
            blocks.bullets.append((0, line[2:].strip()))
        else:
            blocks.paragraphs.append(line)
    _enforce_caps(blocks)
    return blocks
