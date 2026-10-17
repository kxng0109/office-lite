"""CSV bridge: streaming csv to rows and rows to markdown (DRY)."""

from __future__ import annotations

import csv
import io

from .guards import MAX_TABLE_COLS, MAX_TABLE_ROWS, check_size, sanitize_csv_cell, sanitize_md_cell

_ALLOWED_DELIMITERS = (",", ";", "\t", "|", ":")


def read_csv_text(text: str, delimiter: str = ",") -> tuple[list[str], list[list[str]]]:
    if delimiter not in _ALLOWED_DELIMITERS:
        raise ValueError("delimiter must be one of comma, semicolon, tab, pipe, colon")
    try:
        raw = text.encode("utf-8")
    except UnicodeEncodeError as exc:
        raise ValueError("invalid text encoding") from exc
    check_size(raw)
    try:
        reader = csv.reader(io.StringIO(text.lstrip("\ufeff")), delimiter=delimiter)
        header: list[str] = []
        rows: list[list[str]] = []
        for record in reader:
            if not header:
                header = record[:MAX_TABLE_COLS]
                continue
            rows.append(record[:MAX_TABLE_COLS])
            if len(rows) > MAX_TABLE_ROWS:
                raise ValueError("table exceeds row cap")
    except csv.Error as exc:
        raise ValueError("unreadable csv") from exc
    if not header:
        return [], []
    return header, rows


def rows_to_markdown(
    header: list[str], rows: list[list[str]], char_limit: int | None = None
) -> str:
    out = ["| " + " | ".join(sanitize_md_cell(c) for c in header) + " |"]
    out.append("| " + " | ".join("---" for _ in header) + " |")
    omitted = 0
    size = sum(len(line) + 1 for line in out)
    for row in rows:
        padded = (row + [""] * len(header))[: len(header)]
        line = "| " + " | ".join(sanitize_md_cell(c) for c in padded) + " |"
        size += len(line) + 1
        if char_limit is not None and size > char_limit:
            omitted = len(rows) - len(out) + 2
            break
        out.append(line)
    if omitted > 0:
        out.append(f"| ... ({omitted} more rows omitted) |")
    return "\n".join(out) + "\n"


def rows_to_csv_text(header: list[str], rows: list[list[str]]) -> str:
    buf = io.StringIO(newline="")
    writer = csv.writer(buf, lineterminator="\r\n")
    writer.writerow([sanitize_csv_cell(c) for c in header])
    for row in rows:
        padded = (row + [""] * len(header))[: len(header)]
        writer.writerow([sanitize_csv_cell(c) for c in padded])
    return buf.getvalue()
