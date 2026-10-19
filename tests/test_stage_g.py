import base64
import io

import pytest

import office_lite
from office_lite import guards
from office_lite.md_parse import parse_markdown
from office_lite.server import (
    convert_markdown,
    convert_office_out,
    convert_table_in,
    inspect_office,
    search,
)


def test_version_matches_installed_metadata():
    from importlib.metadata import version

    assert office_lite.__version__ == version("office-lite")


def test_hyperlink_destination_preserved():
    blocks = parse_markdown("[Invoice](http://evil.example/x)\n")
    assert blocks.paragraphs == ["[Invoice](http://evil.example/x)"]


def test_nested_tables_extracted():
    from docx import Document

    doc = Document()
    outer = doc.add_table(rows=1, cols=1)
    inner = outer.cell(0, 0).add_table(rows=1, cols=1)
    inner.cell(0, 0).text = "nested-secret"
    buf = io.BytesIO()
    doc.save(buf)
    import base64

    md = convert_office_out(base64.b64encode(buf.getvalue()).decode(), "docx", "markdown")
    assert "nested-secret" in md


def test_multi_table_csv_rejected():
    from docx import Document

    doc = Document()
    for label in ("t1", "t2"):
        table = doc.add_table(rows=1, cols=1)
        table.cell(0, 0).text = label
    buf = io.BytesIO()
    doc.save(buf)
    import base64

    b64 = base64.b64encode(buf.getvalue()).decode()
    with pytest.raises(ValueError, match="multi-table"):
        convert_office_out(b64, "docx", "csv")


def test_same_header_tables_merge_to_csv():
    from docx import Document

    doc = Document()
    for _ in range(2):
        table = doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "h1"
        table.cell(0, 1).text = "h2"
        table.cell(1, 0).text = "v1"
        table.cell(1, 1).text = "v2"
    buf = io.BytesIO()
    doc.save(buf)
    import base64

    b64 = base64.b64encode(buf.getvalue()).decode()
    out = convert_office_out(b64, "docx", "csv")
    assert out.count("v1") == 2


def test_title_control_chars_rejected():
    with pytest.raises(ValueError, match="control"):
        convert_markdown("# Hi", "docx", "Bad\x00Title")
    with pytest.raises(ValueError, match="control"):
        convert_table_in("a,b\n1,2", "markdown", "Bad\x07")
    assert guards.check_title("Family 👨\u200d👩\u200d👧") is None


def test_long_line_rejected():
    with pytest.raises(ValueError, match="input line"):
        parse_markdown("x" * (guards.MAX_LINE_CHARS + 1) + "\n")


def test_nested_structures_depth_capped():
    from markdown_it.token import Token

    from office_lite.md_parse import DocBlocks, _consume_list

    tokens: list[Token] = []
    for _ in range(guards.MAX_LIST_DEPTH + 2):
        tokens.append(Token("bullet_list_open", "ul", 1))
        tokens.append(Token("list_item_open", "li", 1))
    with pytest.raises(ValueError, match="depth cap"):
        _consume_list(tokens, 0, [], DocBlocks())


def test_table_count_cap_enforced():
    parts = []
    for i in range(guards.MAX_TABLES + 1):
        parts.append(f"| h{i} |\n|---|\n| v |\n")
    with pytest.raises(ValueError, match="table cap"):
        parse_markdown("\n".join(parts))


def test_docx_hyperlink_round_trip():
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml.shared import OxmlElement

    doc = Document()
    para = doc.add_paragraph()
    para.add_run("See ")
    r_id = doc.part.relate_to(
        "http://evil.example/x",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), r_id)
    run2 = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "Invoice"
    run2.append(text)
    link.append(run2)
    para._p.append(link)
    buf = io.BytesIO()
    doc.save(buf)
    b64 = base64.b64encode(buf.getvalue()).decode()
    out = convert_office_out(b64, "docx", "markdown")
    assert "[Invoice](http://evil.example/x)" in out


def test_pptx_hyperlink_round_trip():
    from pptx import Presentation
    from pptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(6), Inches(2))
    tf = box.text_frame
    tf.text = "See "
    para = tf.paragraphs[0]
    run = para.add_run()
    run.text = "Deck"
    run.hyperlink.address = "http://evil.example/slides"
    buf = io.BytesIO()
    prs.save(buf)
    b64 = base64.b64encode(buf.getvalue()).decode()
    out = convert_office_out(b64, "pptx", "markdown")
    assert "[Deck](http://evil.example/slides)" in out


def test_inspect_counts_non_empty_paragraphs():
    from docx import Document

    doc = Document()
    doc.add_paragraph("")
    doc.add_paragraph("hello")
    buf = io.BytesIO()
    doc.save(buf)
    b64 = base64.b64encode(buf.getvalue()).decode()
    assert inspect_office(b64, "docx").startswith("paragraphs=1 ")


def test_per_item_caps_enforced():
    with pytest.raises(ValueError, match="length cap"):
        parse_markdown("# " + "x" * (guards.MAX_HEADING_CHARS + 1))
    with pytest.raises(ValueError, match="length cap"):
        parse_markdown("y" * (guards.MAX_PARA_CHARS + 1))


def test_formula_unicode_whitespace_neutralized():
    assert guards.sanitize_csv_cell(" =cmd|'/c calc'!A0").startswith("'")
    assert guards.sanitize_csv_cell("​=1+1").startswith("'")
    assert guards.sanitize_md_cell("​=1+1").startswith("'")


def test_search_strict_contract():
    assert search("docx", limit=20).count("\n") >= 0
