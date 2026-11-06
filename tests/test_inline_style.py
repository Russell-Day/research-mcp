import zipfile

import pytest

from research_mcp.errors import RefusedUnsafe, VerificationFailed
from research_mcp.inline_style import _offline_formatting_snapshot, _validate_request, _verify_formatting


def test_whole_paragraph_is_refused_by_default():
    with pytest.raises(RefusedUnsafe, match="whole-paragraph"):
        _validate_request("A paragraph\r", "A paragraph", allow_whole_paragraph=False)


def test_whole_paragraph_requires_explicit_opt_in():
    assert _validate_request("A paragraph\r", "A paragraph", allow_whole_paragraph=True) == 1


def test_format_verification_rejects_unrequested_font_change():
    before = {(1, 1): ("x", False, False)}
    after = {(1, 1): ("x", True, True)}
    with pytest.raises(VerificationFailed, match="changed formatting"):
        _verify_formatting(before, after, {(1, 1)}, bold=True, italic=None)


def test_format_verification_allows_only_requested_bold_change():
    before = {(1, 1): ("x", False, False)}
    after = {(1, 1): ("x", True, False)}
    _verify_formatting(before, after, {(1, 1)}, bold=True, italic=None)


def test_format_verification_rejects_requested_value_not_reached():
    before = {(1, 1): ("x", False, False)}
    after = {(1, 1): ("x", False, False)}
    with pytest.raises(VerificationFailed):
        _verify_formatting(before, after, {(1, 1)}, bold=True, italic=None)


def test_unbounded_snapshot_preserves_paragraphs_beyond_default_window(monkeypatch, tmp_path):
    captured = {}

    def fake_snapshot(path):
        captured["path"] = path
        return ({"paragraph_index": 41, "text": "x", "bold": False, "italic": False},)

    import research_mcp.inline_style as inline_style
    monkeypatch.setattr(inline_style.style_audit, "snapshot_emphasis", fake_snapshot, raising=False)
    assert _offline_formatting_snapshot(tmp_path / "paper.docx")[(41, 1)] == ("x", False, False)
    assert captured["path"].name == "paper.docx"


def test_real_unbounded_snapshot_includes_late_paragraph(tmp_path):
    body = "".join(f"<w:p><w:r><w:t>paragraph {i}</w:t></w:r></w:p>" for i in range(1, 56))
    path = tmp_path / "late.docx"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>")
        archive.writestr("word/document.xml", f'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr/></w:body></w:document>')
    snapshot = _offline_formatting_snapshot(path)
    assert (55, 1) in snapshot


def test_format_verification_rejects_missing_allowed_character():
    with pytest.raises(VerificationFailed, match="missing"):
        _verify_formatting({}, {}, {(41, 1)}, bold=True, italic=None)
