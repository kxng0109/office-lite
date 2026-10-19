"""Shared guards: validation, truncation, path jail, size caps."""

from __future__ import annotations

import io
import os
import zipfile
from pathlib import Path

from defusedxml.common import DefusedXmlException

MAX_INPUT_BYTES = 10 * 1024 * 1024
MAX_UNCOMPRESSED_BYTES = 50 * 1024 * 1024
MAX_ZIP_RATIO = 100
MAX_ZIP_ENTRIES = 1000
MAX_TABLE_ROWS = 2000
MAX_TABLE_COLS = 100
MAX_SLIDES = 300
MAX_HEADINGS = 500
MAX_TEXT_BLOCKS = 2000
MAX_CODE_BLOCKS = 50
MAX_OUTLINE_ITEMS = 200
MAX_TITLE_CHARS = 200
MAX_HEADING_CHARS = 500
TRUNCATE_CHARS = 25000
TRUNCATE_MARKER = "\n[truncated]"
_XML_SCAN_HEAD = 1024 * 1024


def check_size(data: bytes, limit: int = MAX_INPUT_BYTES) -> None:
    if len(data) > limit:
        raise ValueError("input exceeds size cap")


def audit_zip_bytes(data: bytes) -> None:
    """Enforce zip-entry, expansion, and entity guards on office bytes."""
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            infos = archive.infolist()
            if len(infos) > MAX_ZIP_ENTRIES:
                raise ValueError("archive exceeds entry cap")
            total = sum(item.file_size for item in infos)
            if total > MAX_UNCOMPRESSED_BYTES:
                raise ValueError("archive exceeds expansion cap")
            if len(data) > 0 and total // max(len(data), 1) > MAX_ZIP_RATIO:
                raise ValueError("archive exceeds compression-ratio cap")
            for item in infos:
                name = item.filename.lower()
                if not (name.endswith(".xml") or name.endswith(".rels")):
                    continue
                size = min(item.file_size, _XML_SCAN_HEAD)
                with archive.open(item, "r") as handle:
                    head = handle.read(size).upper()
                if b"<!DOCTYPE" in head or b"<!ENTITY" in head:
                    raise DefusedXmlException("document declares entities")
    except (zipfile.BadZipFile, ValueError, DefusedXmlException):
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
    if candidate != base and base not in candidate.parents:
        raise ValueError("output path escapes base directory")
    return candidate


def ensure_base_dir(base_dir: Path) -> Path:
    base = base_dir.resolve()
    os.makedirs(base, exist_ok=True)
    return base


def sanitize_csv_cell(value: str, for_human_sheet: bool = False) -> str:
    stripped = value.lstrip(" \t") if value else value
    if stripped and stripped[0] in ("=", "+", "-", "@", "\t", "\r", "\n"):
        if for_human_sheet:
            return "\t" + value
        return "'" + value
    return value


def sanitize_md_cell(value: str) -> str:
    """Escape markdown table structure and neutralize formula cells."""
    text = value.replace("\\", "\\\\").replace("|", "\\|")
    text = text.replace("\r\n", "<br/>").replace("\n", "<br/>").replace("\r", "<br/>")
    stripped = text.lstrip(" \t")
    if stripped and stripped[0] in ("=", "+", "-", "@"):
        return "'" + text
    return text
