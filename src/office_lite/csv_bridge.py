"""CSV bridge: streaming csv to rows and rows to markdown (DRY)."""

from __future__ import annotations

import csv
import io

from .guards import MAX_TABLE_COLS, MAX_TABLE_ROWS, check_size, sanitize_csv_cell, sanitize_md_cell

_ALLOWED_DELIMITERS = (",", ";", "\t", "|", ":")

# Explicit bound (not ambient default): the stdlib csv module refuses fields
# larger than this with csv.Error, which read_csv_text converts to ValueError.
# 128 KiB matches CPython's historical default; do not raise it casually.
csv.field_size_limit(128 * 1024)

# Explicit bound (not ambient default): oversized fields fail closed here
# instead of depending on whatever the host interpreter configured.
csv.field_size_limit(128 * 1024)


def read_csv_text(text: str, delimiter: str = ",") -> tuple[list[str], list[list[str]]]:
    if delimiter not in _ALLOWED_DELIMITERS:
        raise ValueError("delimiter must be one of comma, semicolon, tab, pipe, colon")
    try:
        raw = text.encode("utf-8")
    except UnicodeEncodeError as exc:
        raise ValueError("invalid text encoding") from exc
    check_size(raw)
    try:
        reader = csv.reader(
            io.StringIO(text[1:] if text.startswith("\ufeff") else text), delimiter=delimiter
        )
        header: list[str] = []
        rows: list[list[str]] = []
        # NOTE: per-cell length is enforced by the stdlib csv field-size
        # limit; oversized fields raise csv.Error, wrapped below.
        for record in reader:
            if len(record) > MAX_TABLE_COLS:
                raise ValueError("table exceeds column cap")
            if not header:
                header = record
                continue
            rows.append(record)
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
            omitted = len(rows) - (len(out) - 2)
            break
        out.append(line)
    text = "\n".join(out) + "\n"
    if omitted > 0:
        # Marker lives outside the table so strict parsers still see a
        # rectangular grid; humans and machines both see the data loss.
        text += f"\n... ({omitted} more rows omitted)\n"
    return text


def rows_to_csv_text(header: list[str], rows: list[list[str]]) -> str:
    buf = io.StringIO(newline="")
    writer = csv.writer(buf, lineterminator="\r\n")
    writer.writerow([sanitize_csv_cell(c) for c in header])
    for row in rows:
        padded = (row + [""] * len(header))[: len(header)]
        writer.writerow([sanitize_csv_cell(c) for c in padded])
    return buf.getvalue()
