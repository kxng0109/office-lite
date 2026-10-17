"""Shared guards: validation, truncation, path jail, size caps."""

from __future__ import annotations

import io
import re
import zipfile

from defusedxml.common import DefusedXmlException

MAX_INPUT_BYTES = 10 * 1024 * 1024
MAX_UNCOMPRESSED_BYTES = 50 * 1024 * 1024
MAX_ZIP_RATIO = 100
MAX_ZIP_ENTRIES = 1000
MAX_AUDIT_READ_BYTES = 8 * 1024 * 1024
MAX_TABLE_ROWS = 2000
MAX_TABLE_COLS = 100
MAX_CELLS = 50000
MAX_TOTAL_CELLS = 200000
MAX_TABLES = 200
MAX_SLIDES = 300
MAX_GROUP_DEPTH = 32
MAX_LIST_DEPTH = 32
MAX_LINES = 10000
MAX_LINE_CHARS = 1000000
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
# Spaces, tabs, and Unicode whitespace stripped before formula-trigger
# tests so lookalike padding cannot smuggle a trigger past. Members are
# explicit codepoints so every entry stays reviewable in plain ASCII.
_STRIP_LEAD = "".join(
    [" ", "\t"]
    + [
        chr(c)
        for c in (
            0xA0,
            0x1680,
            0x2000,
            0x2001,
            0x2002,
            0x2003,
            0x2004,
            0x2005,
            0x2006,
            0x2007,
            0x2008,
            0x2009,
            0x200A,
            0x2028,
            0x2029,
            0x202F,
            0x205F,
            0x3000,
            0x200B,
            0x200C,
            0x200D,
            0xFEFF,
            0x0085,
            0x000B,
            0x000C,
            0x180E,
            0x2060,
            0x00AD,
            0x3164,
            0xFFA0,
        )
    ]
)


def check_size(data: bytes, limit: int = MAX_INPUT_BYTES) -> None:
    if len(data) > limit:
        raise ValueError("input exceeds size cap")


def _decode_xml_part(data: bytes, name: str) -> str:
    """Decode an XML part for entity scanning, honoring BOM and declaration.

    Tries UTF-8 (with BOM) first, then UTF-16/32 variants. Anything else is
    rejected: an office part we cannot decode is not safe to pass on.
    """
    if data.startswith((b"\xff\xfe\x00\x00", b"\x00\x00\xfe\xff")):
        return data.decode("utf-32")
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        return data.decode("utf-16")
    if data.startswith(b"\xef\xbb\xbf"):
        return data.decode("utf-8-sig")
    # BOM-less UTF-16/32: libxml2-style sniffing on leading NULs, since
    # parsers accept these without a BOM and entities would otherwise hide.
    if data.startswith((b"\x00\x00\x00<", b"<\x00\x00\x00")):
        return data.decode("utf-32-be" if data.startswith(b"\x00") else "utf-32-le")
    if data.startswith(b"\x00<"):
        return data.decode("utf-16-be")
    if data.startswith(b"<\x00"):
        return data.decode("utf-16-le")
    head = data[:200].decode("ascii", errors="ignore")
    if "encoding" in head.lower():
        for encoding in ("utf-8", "utf-16-le", "utf-16-be", "utf-32-le", "utf-32-be"):
            try:
                return data.decode(encoding)
            except (UnicodeDecodeError, LookupError):
                continue
        raise ValueError(f"unreadable xml encoding in {name}")
    try:
        return data.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise ValueError(f"unreadable xml encoding in {name}") from exc


_ENTITY_RE = re.compile(r"<!doctype|<!entity", re.IGNORECASE)


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
            xml_total = sum(
                item.file_size for item in infos if item.filename.lower().endswith(_XML_SCAN_EXTS)
            )
            if xml_total > MAX_AUDIT_READ_BYTES:
                raise ValueError("archive exceeds audit-read cap")
            # Streamed, bounded read: never trust header sizes for allocation,
            # and verify the declared size against actual decompressed bytes.
            # A single entry can never legitimately exceed the audit budget.
            remaining = MAX_AUDIT_READ_BYTES
            for item in infos:
                if not item.filename.lower().endswith(_XML_SCAN_EXTS):
                    continue
                chunks: list[bytes] = []
                got = 0
                with archive.open(item, "r") as handle:
                    while True:
                        if remaining <= 0:
                            raise ValueError("archive exceeds audit-read cap")
                        piece = handle.read(min(65536, remaining))
                        if not piece:
                            break
                        chunks.append(piece)
                        got += len(piece)
                        remaining -= len(piece)
                if got != item.file_size:
                    raise ValueError("archive declares inconsistent sizes")
                text = _decode_xml_part(b"".join(chunks), item.filename)
                if _ENTITY_RE.search(text):
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
    for ch in title:
        code = ord(ch)
        # C0/C1 controls and DEL break XML builders; Cf format chars
        # (e.g. ZWJ in emoji) are legitimate and stay allowed.
        if code < 32 or 0x7F <= code <= 0x9F:
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


# Formula-trigger characters per OWASP CSV Injection (verified live 2026-10-17):
# ASCII = + - @, tab/CR/LF, plus fullwidth variants U+FF1D U+FF0B U+FF0D U+FF20.
# Deliberately NOT included: | % NUL (OWASP lists no such triggers; the DDE
# | separator only matters inside an already-triggered = formula).
_TRIGGER_CHARS = (
    "=",
    "+",
    "-",
    "@",
    "\t",
    "\r",
    "\n",
    "＝",
    "＋",
    "－",
    "＠",
)


def sanitize_csv_cell(value: object) -> str:
    """Neutralize formula cells with a single convention: quote-prefix."""
    text = value if isinstance(value, str) else str(value)
    if text and text[0] not in _STRIP_LEAD and text[0] not in _TRIGGER_CHARS:
        return text
    stripped = text.lstrip(_STRIP_LEAD) if text else text
    if stripped and stripped[0] in _TRIGGER_CHARS:
        return "'" + text
    return text


# URL schemes allowed through to markdown output. Anything else
# (javascript:, data:, file:, vbscript:, blob:) is defanged to plain text.
_SAFE_LINK_SCHEMES = ("http://", "https://", "mailto:")

_DANGEROUS_LINK_RE = None


def _dangerous_link_re():
    global _DANGEROUS_LINK_RE
    if _DANGEROUS_LINK_RE is None:
        import re

        _DANGEROUS_LINK_RE = re.compile(
            r"\[(?P<label>[^\[\]]*)\]\(\s*(?P<scheme>javascript|data|vbscript|file|blob):",
            re.IGNORECASE,
        )
    return _DANGEROUS_LINK_RE


def defang_link_forgery(text: str) -> str:
    """Escape literal text a markdown renderer would turn into a dangerous link.

    Covers text the source never linkified (e.g. markdown-it refuses to linkify
    javascript: URLs, leaving literal ``[click](javascript:...)`` prose that a
    downstream renderer may still linkify).
    """
    return _dangerous_link_re().sub(lambda m: "\\[" + m.group(0)[1:], text)


def safe_link(label: str, url: str) -> str:
    """Render a markdown link, defanging dangerous schemes and forged syntax."""
    text = label.replace("[", "\\[").replace("]", "\\]")
    target = (url or "").strip()
    if not target:
        return text
    low = target.lower()
    if low.startswith("#"):
        return f"[{text}]({target})"
    if not low.startswith(_SAFE_LINK_SCHEMES):
        return text
    if any(ch in target for ch in (" ", "(", ")")):
        return f"[{text}](<{target}>)"
    return f"[{text}]({target})"


def sanitize_md_cell(value: str) -> str:
    """Escape markdown table structure and neutralize formula cells."""
    text = value.replace("\\", "\\\\").replace("|", "\\|")
    text = text.replace("\r\n", "<br/>").replace("\n", "<br/>").replace("\r", "<br/>")
    stripped = text.lstrip(_STRIP_LEAD)
    if stripped and stripped[0] in _TRIGGER_CHARS:
        return "'" + text
    return text
