"""Markdown and text parsing to block stream (DRY for docx/pptx/md)."""

from __future__ import annotations

from dataclasses import dataclass, field

from markdown_it import MarkdownIt

from .guards import MAX_TABLE_COLS, MAX_TABLE_ROWS, check_size

_parser: MarkdownIt | None = None


def get_parser() -> MarkdownIt:
    global _parser
    if _parser is None:
        _parser = MarkdownIt("commonmark").enable("table")
    return _parser


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
        return "".join(c.content for c in token.children if c.content)
    return token.content or ""


def parse_markdown(text: str) -> DocBlocks:
    check_size(text.encode("utf-8", "ignore"))
    blocks = DocBlocks()
    tokens = get_parser().parse(text)
    i = 0
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
            i += 1
        elif tok.type == "list_item_open":
            depth = 0
            j = i + 1
            item_text = ""
            while j < len(tokens) and tokens[j].type != "list_item_close":
                if tokens[j].type == "paragraph_open":
                    inline = tokens[j + 1] if j + 1 < len(tokens) else None
                    if inline is not None:
                        item_text = _inline_text(inline)
                j += 1
            blocks.bullets.append((depth, item_text))
            i = j + 1
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
                        if len(rows) >= MAX_TABLE_ROWS:
                            break
                    current = []
                j += 1
            blocks.tables.append(TableBlock(header=header, rows=rows))
            i = j + 1
        else:
            i += 1
    return blocks


def parse_text(text: str) -> DocBlocks:
    check_size(text.encode("utf-8", "ignore"))
    blocks = DocBlocks()
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        if line.startswith("#"):
            level = len(line) - len(line.lstrip("#"))
            blocks.headings.append((min(level, 6), line.lstrip("#").strip()))
        elif line.startswith(("- ", "* ", "+ ")):
            blocks.bullets.append((0, line[2:].strip()))
        else:
            blocks.paragraphs.append(line)
    return blocks
