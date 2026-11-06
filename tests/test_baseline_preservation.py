"""Word's untouched save is adopted as the verification reference only if it preserves the source."""

from __future__ import annotations

from pathlib import Path
import re
import shutil
import zipfile

import pytest

from research_mcp import edit_batch as B
from research_mcp import editing as E
from research_mcp import fixtures as F
from research_mcp import sections as S
from research_mcp.audit import file_sha256
from research_mcp.edit_models import EditDocumentRequest
from research_mcp.errors import VerificationFailed
from research_mcp.fixtures import create_mixed_run_fixture

CLOSING = F._para(F._run("Plain closing paragraph."))
ANCHORED = (
    '<w:p><w:r><w:t xml:space="preserve">Plain </w:t></w:r>'
    '<w:bookmarkStart w:id="7" w:name="KeepClosing"/><w:commentRangeStart w:id="3"/>'
    "<w:r><w:t>closing</w:t></w:r><w:bookmarkEnd w:id=\"7\"/><w:commentRangeEnd w:id=\"3\"/>"
    '<w:r><w:commentReference w:id="3"/></w:r>'
    '<w:r><w:t xml:space="preserve"> paragraph.</w:t></w:r></w:p>'
)
COMMENTS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:comment w:id="{id}" w:author="Reviewer" w:initials="R"><w:p><w:r><w:t>Check this.</w:t></w:r></w:p>'
    "</w:comment></w:comments>"
)


def _package(source: Path, target: Path, mutate=None, comment_id: str = "3") -> Path:
    with zipfile.ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    xml = parts["word/document.xml"].decode("utf-8")
    if mutate is not None:
        changed = mutate(xml)
        assert changed != xml
        xml = changed
    parts["word/document.xml"] = xml.encode("utf-8")
    parts["word/comments.xml"] = COMMENTS.format(id=comment_id).encode("utf-8")
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return target


@pytest.fixture
def anchored(tmp_path: Path) -> Path:
    mixed = create_mixed_run_fixture(tmp_path / "mixed.docx")
    return _package(mixed, tmp_path / "anchored.docx", lambda xml: xml.replace(CLOSING, ANCHORED))


def _variant(anchored: Path, tmp_path: Path, name: str, mutate=None, comment_id: str = "3") -> Path:
    return _package(anchored, tmp_path / f"{name}.docx", mutate, comment_id)


def test_inventory_records_every_anchor_with_its_text_extent(anchored):
    inventory = S.anchor_inventory(anchored)
    assert inventory["bookmarks"] == [("KeepClosing", [((22, 6), (22, 13), None, None)])]
    kinds = sorted(item[0] for item in inventory["comments"])
    assert kinds == ["range", "reference"]
    assert [field[:2] for field in inventory["fields"]] == [("complex", 'QUOTE "Example 2024"')]
    assert len(inventory["content controls"]) == 1 and len(inventory["hyperlinks"]) == 1


def test_word_save_normalization_is_tolerated(anchored, tmp_path):
    def renumber_and_resplit(xml: str) -> str:
        xml = xml.replace('w:id="7"', 'w:id="0"').replace('<w:id w:val="4242"/>', '<w:id w:val="-77"/>')
        xml = xml.replace('w:commentRangeStart w:id="3"', 'w:commentRangeStart w:id="9"')
        xml = xml.replace('w:commentRangeEnd w:id="3"', 'w:commentRangeEnd w:id="9"')
        xml = xml.replace('w:commentReference w:id="3"', 'w:commentReference w:id="9"')
        # Equivalent runs split differently, an rsid, and Word's own _GoBack bookmark.
        xml = xml.replace(F._run("Jordan Example", "b", "sz=36"),
                          '<w:r w:rsidR="00AB12CD">' + F._run("Jordan", "b", "sz=36")[5:]
                          + F._run(" Example", "b", "sz=36"))
        return xml.replace("<w:body><w:p>", '<w:body><w:p><w:bookmarkStart w:id="99" w:name="_GoBack"/><w:bookmarkEnd w:id="99"/>', 1)

    baseline = _variant(anchored, tmp_path, "normalized", renumber_and_resplit, comment_id="9")
    assert S.baseline_preservation_errors(anchored, baseline) == []


@pytest.mark.parametrize(
    ("name", "mutate", "message"),
    [
        ("collapsed", lambda xml: xml.replace('<w:bookmarkEnd w:id="7"/>', "").replace(
            '<w:bookmarkStart w:id="7" w:name="KeepClosing"/>',
            '<w:bookmarkStart w:id="7" w:name="KeepClosing"/><w:bookmarkEnd w:id="7"/>'),
         r"bookmarks: .*KeepClosing.*\(22, 13\).* became .*\(22, 6\), \(22, 6\)"),
        ("lost", lambda xml: xml.replace('<w:bookmarkEnd w:id="7"/>', "").replace(
            '<w:bookmarkStart w:id="7" w:name="KeepClosing"/>', ""), r"bookmarks: 1 -> 0 \(lost"),
        ("unbolded", lambda xml: xml.replace(F._run("clinical", "b", "i"), F._run("clinical", "i")),
         r"paragraph 2 character 29 \('c'\) formatting \[b i\] became \[i\]"),
        ("comment_moved", lambda xml: xml.replace('<w:commentRangeEnd w:id="3"/>', "").replace(
            '<w:r><w:t xml:space="preserve"> paragraph.</w:t></w:r></w:p>',
            '<w:r><w:t xml:space="preserve"> paragraph.</w:t></w:r><w:commentRangeEnd w:id="3"/></w:p>'),
         r"comments: "),
        ("field_changed", lambda xml: xml.replace('QUOTE "Example 2024"', 'QUOTE "Example 2025"'), r"fields: "),
        ("control_changed", lambda xml: xml.replace('<w:alias w:val="Venue"/>', '<w:alias w:val="Place"/>'),
         r"content controls: "),
        ("link_changed", lambda xml: xml.replace('w:anchor="Top"', 'w:anchor="Bottom"'), r"hyperlinks: "),
    ],
)
def test_baseline_damage_is_reported(anchored, tmp_path, name, mutate, message):
    baseline = _variant(anchored, tmp_path, name, mutate)
    problems = S.baseline_preservation_errors(anchored, baseline)
    assert problems and any(re.search(message, item) for item in problems), problems


class _DamagingWordSave:
    """Stands in for Word: the first (untouched) save applies `damage`; later saves copy."""

    damage = staticmethod(lambda xml: xml)

    def __init__(self, path, **kwargs):
        self.path = Path(path)
        self.document = self
        self.first = True

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False

    def save_as(self, target):
        target = Path(target)
        shutil.copyfile(self.path, target)
        self.path = target
        if self.first:
            self.first = False
            with zipfile.ZipFile(target) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
            parts["word/document.xml"] = self.damage(parts["word/document.xml"].decode("utf-8")).encode("utf-8")
            with zipfile.ZipFile(target, "w") as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
        return target


def _rename_title(document, plan, **kwargs):
    path = Path(document.path)
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts["word/document.xml"] = parts["word/document.xml"].replace(b"Jordan Example", b"Jordyn Example")
    with zipfile.ZipFile(path, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


COLLAPSE = staticmethod(lambda xml: xml.replace('<w:bookmarkEnd w:id="7"/>', "").replace(
    '<w:bookmarkStart w:id="7" w:name="KeepClosing"/>',
    '<w:bookmarkStart w:id="7" w:name="KeepClosing"/><w:bookmarkEnd w:id="7"/>'))
UNBOLD = staticmethod(lambda xml: xml.replace(F._run("clinical", "b", "i"), F._run("clinical", "i")))


@pytest.mark.parametrize("damage", [COLLAPSE, UNBOLD], ids=["collapsed_bookmark", "unrelated_bold"])
def test_damaging_baseline_is_not_adopted(anchored, tmp_path, monkeypatch, damage):
    monkeypatch.setattr(_DamagingWordSave, "damage", damage)
    monkeypatch.setattr(E, "IsolatedWordDocument", _DamagingWordSave)
    monkeypatch.setattr(E, "_execute_plan", _rename_title)
    with pytest.raises(VerificationFailed, match="did not preserve it"):
        E.replace_paragraph_to_version(anchored, tmp_path / "out.docx", paragraph_index=1, text="Jordyn Example")


def test_clean_baseline_is_adopted(anchored, tmp_path, monkeypatch):
    monkeypatch.setattr(E, "IsolatedWordDocument", _DamagingWordSave)
    monkeypatch.setattr(E, "_execute_plan", _rename_title)
    result = E.replace_paragraph_to_version(anchored, tmp_path / "out.docx", paragraph_index=1, text="Jordyn Example")
    assert result["verification"]["compared_against"] == "word_baseline"


def test_batch_with_damaging_baseline_publishes_nothing(anchored, tmp_path, monkeypatch):
    monkeypatch.setattr(_DamagingWordSave, "damage", COLLAPSE)
    monkeypatch.setattr(E, "IsolatedWordDocument", _DamagingWordSave)
    monkeypatch.setattr(E, "_execute_plan", _rename_title)
    monkeypatch.setattr(B, "_apply_operation", B._apply_operation_local)
    output = tmp_path / "published.docx"
    request = EditDocumentRequest(
        source=str(anchored), source_sha256=file_sha256(anchored), output=str(output),
        operations=[{"op_id": "title", "type": "replace_paragraph", "paragraph_index": 1, "text": "Jordyn Example"}])
    result = B.edit_document(request)
    assert result.status == "no_changes"
    assert result.operations[0].error.code == "VerificationFailed"
    assert "did not preserve it" in result.operations[0].error.message
    assert not output.exists() and result.undo_token is None
    assert file_sha256(anchored) == request.source_sha256
