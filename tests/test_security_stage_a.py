import io
import zipfile

import pytest
from defusedxml.common import DefusedXmlException

from office_lite import guards
from office_lite.csv_bridge import rows_to_markdown


def _office_zip(name: str, body: str) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        archive.writestr(name, body)
    return buf.getvalue()


def test_audit_rejects_many_entries():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as archive:
        for i in range(guards.MAX_ZIP_ENTRIES + 1):
            archive.writestr(f"word/p{i}.xml", "<a/>")
    with pytest.raises(ValueError, match="entry cap"):
        guards.audit_zip_bytes(buf.getvalue())


def test_audit_rejects_entities():
    data = _office_zip(
        "word/document.xml",
        '<?xml version="1.0"?><!DOCTYPE r [<!ENTITY x "y">]><r>&x;</r>',
    )
    with pytest.raises(DefusedXmlException):
        guards.audit_zip_bytes(data)


def test_audit_accepts_clean_office_like_zip():
    data = _office_zip("word/document.xml", "<w:document><w:p><w:t>hi</w:t></w:p></w:document>")
    guards.audit_zip_bytes(data)


def test_caps_enforced_on_blocks():
    from office_lite.md_parse import DocBlocks, parse_markdown

    many = "\n\n".join(f"# H{i}" for i in range(guards.MAX_HEADINGS + 1))
    with pytest.raises(ValueError, match="heading cap"):
        parse_markdown(many)
    blocks = DocBlocks(headings=[(1, "t")] * (guards.MAX_SLIDES + 1))
    from office_lite.pptx_build import build_pptx

    with pytest.raises(ValueError, match="slide cap"):
        build_pptx(blocks)


def test_markdown_cells_sanitized_and_escaped():
    md = rows_to_markdown(["h"], [["=1+1"], ["a|b"], ["x\ny"]])
    assert "'=1+1" in md
    assert "a\\|b" in md
    assert "<br/>" in md


def test_b64_validate_rejects_whitespace_smuggle():
    from office_lite.server import _decode_office_b64

    with pytest.raises(ValueError, match="base64"):
        _decode_office_b64("!!!not-b64!!!")
    with pytest.raises(ValueError, match="size cap"):
        _decode_office_b64("A" * (guards.MAX_INPUT_BYTES * 4 // 3 + 2048))
