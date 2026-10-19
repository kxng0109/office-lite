import base64
import io
import zipfile

import pytest

from office_lite import guards
from office_lite import server as S
from office_lite.docx_build import _new_doc
from office_lite.md_parse import get_parser
from office_lite.pptx_build import _new_deck


def _b64(data: bytes) -> str:
    return base64.b64encode(data).decode()


def test_audit_entry_cap():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as archive:
        for i in range(guards.MAX_ZIP_ENTRIES + 1):
            archive.writestr(f"word/p{i}.xml", "<a/>")
    with pytest.raises(ValueError, match="entry cap"):
        guards.audit_zip_bytes(buf.getvalue())


def test_audit_ratio_cap_with_deflate_bomb():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", "<a>" + "0" * 30000 + "</a>")
    with pytest.raises(ValueError, match="ratio cap"):
        guards.audit_zip_bytes(buf.getvalue())


def test_audit_read_budget_cap():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as archive:
        for i in range(3):
            archive.writestr(f"word/p{i}.xml", "<a>" + "x" * (3 * 1024 * 1024) + "</a>")
    with pytest.raises(ValueError, match="audit-read cap"):
        guards.audit_zip_bytes(buf.getvalue())


def test_open_helpers_reject_broken_packages():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
    raw = buf.getvalue()
    with pytest.raises(ValueError, match="invalid docx"):
        S._open_docx(raw)
    with pytest.raises(ValueError, match="invalid pptx"):
        S._open_pptx(raw)
    buf2 = io.BytesIO()
    with zipfile.ZipFile(buf2, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        archive.writestr("word/document.xml", "not <xml at all <<<")
    with pytest.raises(ValueError, match="invalid docx"):
        S._open_docx(buf2.getvalue())


def test_b64_input_cap():
    with pytest.raises(ValueError, match="size cap"):
        S._decode_office_b64("A" * (guards.MAX_INPUT_BYTES * 4 // 3 + 2048))


def test_strict_input_types():
    with pytest.raises(ValueError, match="must be a string"):
        S.convert_markdown(123, "docx")  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="must be a string"):
        S.convert_table_in(None, "markdown")  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="must be a string"):
        S._decode_office_b64(123)  # type: ignore[arg-type]


def test_template_cache_reused_and_independent():
    first = _new_doc()
    second = _new_doc()
    assert first is not second
    first_pptx = _new_deck()
    second_pptx = _new_deck()
    assert first_pptx is not second_pptx


def test_walk_depth_caps_with_fakes():
    from types import SimpleNamespace

    def nest_group(depth):
        if depth == 0:
            return SimpleNamespace(shape_type=0, shapes=[])
        return SimpleNamespace(shape_type=6, shapes=[nest_group(depth - 1)])

    from pptx.enum.shapes import MSO_SHAPE_TYPE

    assert MSO_SHAPE_TYPE.GROUP == 6
    with pytest.raises(ValueError, match="depth cap"):
        list(S._walk_shapes([nest_group(guards.MAX_GROUP_DEPTH + 2)]))

    def nest_table(depth):
        if depth == 0:
            return SimpleNamespace(rows=[])
        cell = SimpleNamespace(tables=[nest_table(depth - 1)])
        return SimpleNamespace(rows=[SimpleNamespace(cells=[cell])])

    with pytest.raises(ValueError, match="depth cap"):
        list(S._walk_tables([nest_table(guards.MAX_GROUP_DEPTH + 2)]))


def test_get_parser_shared():
    assert get_parser() is get_parser()


def test_main_failure_exits_nonzero(monkeypatch):
    import office_lite.server as srv

    def boom():
        raise RuntimeError("loop dead")

    monkeypatch.setattr(srv.server, "run", boom)
    with pytest.raises(SystemExit) as exc:
        srv.main()
    assert exc.value.code == 1
