import pytest

from office_lite.md_parse import parse_markdown, parse_text
from office_lite.server import (
    convert_markdown,
    convert_table_in,
    create_from_outline,
    inspect_office,
    search,
)


def test_flavor_allowlist():
    with pytest.raises(ValueError, match="flavor"):
        convert_markdown("# Hi", "docx", "T", "md")
    assert convert_markdown("# Hi", "docx", "T", "text")


def test_delimiter_validation():
    with pytest.raises(ValueError, match="delimiter"):
        convert_table_in("a,b", "markdown", "T", "")
    with pytest.raises(ValueError, match="delimiter"):
        convert_table_in("a,b", "markdown", "T", ";;")
    with pytest.raises(ValueError, match="delimiter"):
        convert_table_in("a,b", "markdown", "T", "\n")


def test_title_and_outline_validation():
    with pytest.raises(ValueError, match="title"):
        convert_markdown("# Hi", "docx", "T" * 201)
    with pytest.raises(ValueError, match="headings"):
        create_from_outline("not-a-list", "docx")  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="non-empty"):
        create_from_outline(["ok", 123], "docx")  # type: ignore[list-item]
    with pytest.raises(ValueError, match="item cap"):
        create_from_outline([f"h{i}" for i in range(201)], "docx")
    with pytest.raises(ValueError, match="source"):
        inspect_office("e30=", "xlsx")


def test_search_contract():
    assert "convert_markdown" in search("MARKDOWN")
    with pytest.raises(ValueError, match="integer"):
        search("md", limit="10")  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="integer"):
        search("md", limit=True)  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="string"):
        search(123)  # type: ignore[arg-type]


def test_ordered_and_nested_lists():
    blocks = parse_markdown("1. first\n2. second\n\n- a\n  - b\n")
    texts = [t for _, t in blocks.bullets]
    assert texts[0].startswith("1. ")
    assert texts[1].startswith("2. ")
    assert any(depth == 1 for depth, _ in blocks.bullets)


def test_hash_tag_is_not_heading():
    blocks = parse_text("#tag prose")
    assert blocks.headings == []
    assert blocks.paragraphs == ["#tag prose"]
    assert parse_text("#").headings == []


def test_strict_encoding_rejects_surrogates():
    with pytest.raises(ValueError, match="encoding"):
        convert_markdown("bad \ud800 char", "docx")
    with pytest.raises(ValueError, match="encoding"):
        convert_table_in("a\ud800,b", "markdown")
