"""Shared guards: validation, truncation, path jail, size caps."""

from __future__ import annotations

import io
import os
import unicodedata
import zipfile
from pathlib import Path

from defusedxml.common import DefusedXmlException

MAX_INPUT_BYTES = 10 * 1024 * 1024
MAX_UNCOMPRESSED_BYTES = 50 * 1024 * 1024
MAX_ZIP_RATIO = 100
MAX_ZIP_ENTRIES = 1000
MAX_AUDIT_READ_BYTES = 8 * 1024 * 1024
MAX_TABLE_ROWS = 2000
MAX_TABLE_COLS = 100
MAX_CELLS = 50000
MAX_SLIDES = 300
MAX_GROUP_DEPTH = 32
MAX_LINES = 10000
MAX_HEADINGS = 500
MAX_TEXT_BLOCKS = 2000
MAX_CODE_BLOCKS = 50
MAX_OUTLINE_ITEMS = 200
MAX_TITLE_CHARS = 200
MAX_HEADING_CHARS = 500
MAX_PARA_CHARS = 100000
TRUNCATE_CHARS = 25000
TRUNCATE_MARKER = "\n[truncated]"
_XML_SCAN_EXTS = (".xml", ".rels")
# Spaces, tabs, and Unicode whitespace (NBSP, OGHAM, EM/EN quads, hair/thin
# spaces, ideographic space, zero-width space/joiners, BOM) stripped before
# formula-trigger tests so lookalike padding cannot smuggle a trigger past.
_STRIP_LEAD = " \t               　​‌‍﻿"


def check_size(data: bytes, limit: int = MAX_INPUT_BYTES) -> None:
    if len(data) > limit:
        raise ValueError("input exceeds size cap")


def audit_zip_bytes(data: bytes) -> None:
    """Enforce zip-entry, expansion, ratio, and entity guards on office bytes."""
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            infos = archive.infolist()
            if len(infos) > MAX_ZIP_ENTRIES:
                raise ValueError("archive exceeds entry cap")
            total = sum(item.file_size for item in infos)
            if total > MAX_UNCOMPRESSED_BYTES:
                raise ValueError("archive exceeds expansion cap")
            if len(data) > 0 and total > len(data) * MAX_ZIP_RATIO:
                raise ValueError("archive exceeds compression-ratio cap")
            budget = MAX_AUDIT_READ_BYTES
            for item in infos:
                if not item.filename.lower().endswith(_XML_SCAN_EXTS):
                    continue
                if budget <= 0:
                    raise ValueError("archive exceeds audit-read cap")
                with archive.open(item, "r") as handle:
                    head = handle.read(min(item.file_size, budget)).upper()
                budget -= len(head)
                if b"<!DOCTYPE" in head or b"<!ENTITY" in head:
                    raise DefusedXmlException("document declares entities")
    except zipfile.BadZipFile as exc:
        raise ValueError("unreadable archive") from exc
    except (ValueError, DefusedXmlException):
        raise
    except Exception as exc:
        raise ValueError("unreadable archive") from exc


def truncate_text(text: str, limit: int = TRUNCATE_CHARS) -> str:
    if len(text) <= limit:
        return text
    return text[:limit] + TRUNCATE_MARKER


def check_title(title: str, limit: int = MAX_TITLE_CHARS) -> None:
    if not isinstance(title, str) or not title or len(title) > limit:
        raise ValueError("title must be 1-200 characters")
    if any(unicodedata.category(ch).startswith("C") for ch in title):
        raise ValueError("title must not contain control characters")


def check_outline(headings: object, limit: int = MAX_OUTLINE_ITEMS) -> list[str]:
    if not isinstance(headings, list):
        raise ValueError("headings must be a list of strings")
    if len(headings) > limit:
        raise ValueError("outline exceeds item cap")
    cleaned: list[str] = []
    for item in headings:
        if not isinstance(item, str) or not item.strip():
            raise ValueError("outline items must be non-empty strings")
        if len(item) > MAX_HEADING_CHARS:
            raise ValueError("outline item exceeds length cap")
        cleaned.append(item.strip())
    return cleaned


def safe_output_path(base_dir: Path, name: str) -> Path:
    base = base_dir.resolve()
    candidate = (base / name).resolve()
    if candidate == base or base not in candidate.parents:
        raise ValueError("output path escapes base directory")
    return candidate


def ensure_base_dir(base_dir: Path) -> Path:
    base = base_dir.resolve()
    os.makedirs(base, exist_ok=True)
    return base


def sanitize_csv_cell(value: object) -> str:
    """Neutralize formula cells with a single convention: quote-prefix."""
    text = value if isinstance(value, str) else str(value)
    stripped = text.lstrip(_STRIP_LEAD) if text else text
    if stripped and stripped[0] in ("=", "+", "-", "@", "\t", "\r", "\n"):
        return "'" + text
    return text


def sanitize_md_cell(value: str) -> str:
    """Escape markdown table structure and neutralize formula cells."""
    text = value.replace("\\", "\\\\").replace("|", "\\|")
    text = text.replace("\r\n", "<br/>").replace("\n", "<br/>").replace("\r", "<br/>")
    stripped = text.lstrip(_STRIP_LEAD)
    if stripped and stripped[0] in ("=", "+", "-", "@"):
        return "'" + text
    return text
