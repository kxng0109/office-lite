"""CSV bridge: streaming csv to rows and rows to markdown (DRY)."""

from __future__ import annotations

import csv
import io

from .guards import MAX_TABLE_COLS, MAX_TABLE_ROWS, check_size, sanitize_csv_cell


def read_csv_text(text: str, delimiter: str = ",") -> tuple[list[str], list[list[str]]]:
    check_size(text.encode("utf-8", "ignore"))
    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    rows = [r[:MAX_TABLE_COLS] for r in reader][: MAX_TABLE_ROWS + 1]
    if not rows:
        return [], []
    return rows[0], rows[1:]


def rows_to_markdown(header: list[str], rows: list[list[str]]) -> str:
    def esc(cell: str) -> str:
        return cell.replace("|", "\\|")

    out = ["| " + " | ".join(esc(c) for c in header) + " |"]
    out.append("| " + " | ".join("---" for _ in header) + " |")
    for row in rows:
        padded = (row + [""] * len(header))[: len(header)]
        out.append("| " + " | ".join(esc(c) for c in padded) + " |")
    return "\n".join(out) + "\n"


def rows_to_csv_text(
    header: list[str], rows: list[list[str]], for_human_sheet: bool = False
) -> str:
    buf = io.StringIO(newline="")
    writer = csv.writer(buf, lineterminator="\r\n")
    writer.writerow([sanitize_csv_cell(c, for_human_sheet) for c in header])
    for row in rows:
        padded = (row + [""] * len(header))[: len(header)]
        writer.writerow([sanitize_csv_cell(c, for_human_sheet) for c in padded])
    return buf.getvalue()
