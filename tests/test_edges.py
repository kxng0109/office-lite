import base64

import pytest

from office_lite import guards
from office_lite.csv_bridge import read_csv_text
from office_lite.docx_build import build_docx, build_docx_from_rows
from office_lite.md_parse import TableBlock, parse_markdown
from office_lite.pptx_build import build_pptx, build_pptx_from_rows
from office_lite.server import (
    convert_markdown,
    convert_office_out,
    convert_table_in,
    create_from_outline,
    inspect_office,
)


def test_guards_edges(tmp_path):
    guards.ensure_base_dir(tmp_path / "out")
    assert (tmp_path / "out").exists()
    with pytest.raises(ValueError):
        guards.check_size(b"x" * (guards.MAX_INPUT_BYTES + 1))
    assert guards.sanitize_csv_cell("plain") == "plain"


def test_markdown_table_and_text_flavor():
    blocks = parse_markdown("| a | b |\n|---|---|\n| 1 | 2 |\n| 3 | 4 |\n")
    assert len(blocks.tables) == 1
    assert blocks.tables[0].header == ["a", "b"]
    docx = build_docx(blocks, "")
    assert docx[:2] == b"PK"
    empty = build_docx(parse_markdown(""), "")
    assert empty[:2] == b"PK"


def test_table_limits_and_empty():
    header = [f"c{i}" for i in range(5)]
    rows = [["1"] * 5]
    assert build_docx_from_rows([], [])[:2] == b"PK"
    assert build_pptx_from_rows([], [])[:2] == b"PK"
    assert build_pptx(parse_markdown("# T\n```\ncode\n```\n"), "T")[:2] == b"PK"
    assert build_docx_from_rows(header, rows, "")[:2] == b"PK"


def test_server_error_and_csv_text_paths():
    with pytest.raises(ValueError):
        convert_markdown("x", "xlsx")
    with pytest.raises(ValueError):
        convert_table_in("a", "xlsx")
    with pytest.raises(ValueError):
        convert_office_out(base64.b64encode(b"PKxx").decode(), "xlsx", "markdown")
    txt_b64 = convert_markdown("plain line", "docx", "T", "text")
    assert "plain" in convert_office_out(txt_b64, "docx", "text")
    assert convert_office_out(txt_b64, "docx", "csv") == ""
    header, rows = read_csv_text("a;b\n1;2\n", ";")
    assert header == ["a", "b"]
    md = convert_table_in("a,b\n1,2", "markdown")
    assert "|" in md
    csv_out = convert_table_in("a,b\n1,2", "csv")
    assert "a" in csv_out
    pptx_b64 = convert_table_in("a,b\n1,2", "pptx", "T")
    assert "Slide" in convert_office_out(pptx_b64, "pptx", "text")
    assert "Slide" in convert_office_out(pptx_b64, "pptx", "markdown")
    created = create_from_outline(["A"], "pptx", "T")
    assert base64.b64decode(created)[:2] == b"PK"
    with pytest.raises(ValueError):
        create_from_outline(["A"], "xlsx")
    assert "bytes=" in inspect_office(txt_b64, "docx")
    empty_table = TableBlock(header=[], rows=[])
    assert build_docx_from_rows(empty_table.header, empty_table.rows)[:2] == b"PK"
