"""Shared guards: validation, truncation, path jail, size caps."""

from __future__ import annotations

import os
from pathlib import Path

MAX_INPUT_BYTES = 10 * 1024 * 1024
MAX_ZIP_ENTRIES = 1000
MAX_TABLE_ROWS = 2000
MAX_TABLE_COLS = 100
TRUNCATE_CHARS = 25000
TRUNCATE_MARKER = "\n[truncated]"


def check_size(data: bytes, limit: int = MAX_INPUT_BYTES) -> None:
    if len(data) > limit:
        raise ValueError("input exceeds size cap")


def truncate_text(text: str, limit: int = TRUNCATE_CHARS) -> str:
    if len(text) <= limit:
        return text
    return text[:limit] + TRUNCATE_MARKER


def safe_output_path(base_dir: Path, name: str) -> Path:
    base = base_dir.resolve()
    candidate = (base / name).resolve()
    if candidate != base and base not in candidate.parents:
        raise ValueError("output path escapes base directory")
    return candidate


def ensure_base_dir(base_dir: Path) -> Path:
    base = base_dir.resolve()
    os.makedirs(base, exist_ok=True)
    return base


def sanitize_csv_cell(value: str, for_human_sheet: bool = False) -> str:
    if value and value[0] in ("=", "+", "-", "@", "\t", "\r", "\n"):
        if for_human_sheet:
            return "\t" + value
        return "'" + value
    return value
