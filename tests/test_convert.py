import base64

from office_lite import guards
from office_lite.csv_bridge import read_csv_text, rows_to_csv_text, rows_to_markdown
from office_lite.docx_build import build_docx, build_docx_from_rows
from office_lite.md_parse import parse_markdown, parse_text
from office_lite.pptx_build import build_pptx, build_pptx_from_rows
from office_lite.server import (
    convert_markdown,
    convert_office_out,
    convert_table_in,
    create_from_outline,
    inspect_office,
    search,
)


def test_guards_truncate_and_jail(tmp_path):
    long = "x" * (guards.TRUNCATE_CHARS + 10)
    assert guards.TRUNCATE_MARKER in guards.truncate_text(long)
    base = tmp_path
    assert guards.safe_output_path(base, "a.docx").name == "a.docx"
    try:
        guards.safe_output_path(base, "../evil.docx")
        raise AssertionError("should jail")
    except ValueError:
        pass
    assert guards.sanitize_csv_cell("=1+1").startswith("'")
    assert guards.sanitize_csv_cell("plain") == "plain"


def test_parse_and_csv():
    blocks = parse_markdown("# T\n\nHello\n\n- a\n\n```\ncode\n```\n")
    assert blocks.headings and blocks.paragraphs and blocks.bullets and blocks.codes
    text_blocks = parse_text("# H\n- x\nplain")
    assert text_blocks.headings and text_blocks.bullets
    header, rows = read_csv_text("a,b\n1,2\n")
    assert header == ["a", "b"]
    md = rows_to_markdown(header, rows)
    assert "| a | b |" in md
    csv_text = rows_to_csv_text(header, [["=1+1", "x"]])
    assert "'=1+1" in csv_text


def test_build_docx_pptx_bytes():
    blocks = parse_markdown("# Deck\n\n- one\n\n| a | b |\n|---|---|\n| 1 | 2 |\n")
    docx = build_docx(blocks, "T")
    assert docx[:2] == b"PK"
    pptx = build_pptx(blocks, "T")
    assert pptx[:2] == b"PK"
    docx2 = build_docx_from_rows(["a"], [["1"]])
    assert docx2[:2] == b"PK"
    pptx2 = build_pptx_from_rows(["a"], [["1"]])
    assert pptx2[:2] == b"PK"


def test_server_tools_roundtrip():
    docx_b64 = convert_markdown("# Hi\n\n- a", "docx", "T", "markdown")
    assert inspect_office(docx_b64, "docx").startswith("paragraphs=")
    md = convert_office_out(docx_b64, "docx", "markdown")
    assert "Hi" in md
    pptx_b64 = convert_table_in("a,b\n1,2", "pptx")
    assert inspect_office(pptx_b64, "pptx").startswith("slides=")
    out_csv = convert_office_out(pptx_b64, "pptx", "csv")
    assert "a" in out_csv
    created = create_from_outline(["One", "Two"], "docx")
    assert base64.b64decode(created)[:2] == b"PK"
    assert "Convert markdown" in search("markdown")
    try:
        search("")
        raise AssertionError("should reject")
    except ValueError:
        pass
