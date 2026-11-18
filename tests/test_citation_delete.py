"""Plan 831 S4: delete_citation removes one whole Zotero citation field as a tracked deletion.

Synthetic docx (no Word, no Zotero). Expectations are read back with this file's own regex/ElementTree
reader and through the MCP's own accept/reject (`revisions.resolve_revisions`), the path the user's Accept
or Reject in the MCP takes.
"""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path

import pytest

from research_mcp.citation_delete import delete_citation
from research_mcp.errors import RefusedUnsafe
from research_mcp.revisions import list_revisions, resolve_revisions

URI = "http://zotero.org/users/1/items/"
NUMBERS = {"K1": 1, "K2": 2, "K3": 3}
DATE = "2026-11-01T10:00:00Z"
INS = f'<w:ins w:id="101" w:author="Sam" w:date="{DATE}"><w:r><w:t xml:space="preserve"> added</w:t></w:r></w:ins>'
DEL = f'<w:del w:id="102" w:author="Reviewer" w:date="{DATE}"><w:r><w:delText>old </w:delText></w:r></w:del>'


def _instr(cid: str, keys: list[str], formatted: str) -> str:
    items = [{"id": NUMBERS[k], "uris": [URI + k], "itemData": {"id": NUMBERS[k], "title": f"Title {k}"}}
             for k in keys]
    payload = {"citationID": cid, "properties": {"formattedCitation": formatted, "plainCitation": formatted,
                                                 "noteIndex": 0}, "citationItems": items}
    return (" ADDIN ZOTERO_ITEM CSL_CITATION " + json.dumps(payload, separators=(",", ":")) + " "
            ).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def field(cid: str, keys: list[str], rendered: str, rpr: str = "") -> str:
    return (f'<w:r>{rpr}<w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r>{rpr}<w:instrText xml:space="preserve">{_instr(cid, keys, rendered)}</w:instrText></w:r>'
            f'<w:r>{rpr}<w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r>{rpr}<w:t>{rendered}</w:t></w:r><w:r>{rpr}<w:fldChar w:fldCharType="end"/></w:r>')


def t(text: str, rpr: str = "") -> str:
    return f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'


BIBL = ('<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} '
        'CSL_BIBLIOGRAPHY </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:t>1. Title K1</w:t></w:r></w:p><w:p><w:r><w:t>2. Title K2</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>3. Title K3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')


def make_docx(path: Path, *paragraphs: str, bibl: bool = True) -> Path:
    body = "".join(f"<w:p>{p}</w:p>" for p in paragraphs) + (BIBL if bibl else "") + "<w:sectPr/>"
    document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                f"<w:body>{body}</w:body></w:document>")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml",
                         '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                         '<Default Extension="xml" ContentType="application/xml"/></Types>')
        archive.writestr("word/document.xml", document)
    return path


def doc_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        return archive.read("word/document.xml").decode("utf-8")


def para_texts(path: Path) -> list[str]:
    """Visible text per paragraph as Word shows it with markup off (w:t only, field codes excluded)."""
    return [result["accepted_text"] for result in list_revisions_all(path)]


def list_revisions_all(path: Path) -> list[dict]:
    xml = doc_xml(path)
    count = len(re.findall(r"<w:p[ >]", xml))
    return list_revisions(path, paragraph_index=1, through_paragraph_index=count, limit=1000)["paragraphs"]


def instr_payloads(path: Path) -> list[str]:
    """Live (not deleted) ZOTERO_ITEM field codes, in order."""
    return [m for m in re.findall(r"<w:instrText[^>]*>([^<]*)</w:instrText>", doc_xml(path)) if "ZOTERO_ITEM" in m]


def rev_elements(path: Path) -> list[str]:
    return re.findall(r"<w:(?:ins|del)\b[^>]*>.*?</w:(?:ins|del)>", doc_xml(path), re.S)


LEGEND = t("Figure 2 legend ") + field("cX", ["K1"], "22") + t(".")
OTHER = t("Body cites ") + field("cB", ["K2"], "(2)") + t(" here.") + INS
ITEMS = t("Both ") + field("cM", ["K1", "K3"], "(1,3)") + t(" and ") + field("cK", ["K1"], "(1)") + DEL + t("x")


@pytest.fixture
def doc(tmp_path: Path) -> Path:
    return make_docx(tmp_path / "in.docx", LEGEND, OTHER, ITEMS)


def test_tracked_delete_of_stray_legend_citation(doc, tmp_path):
    """chk_2b7494a104d140fa: the stray "22" leaves as one tracked deletion by Claude, with its space."""
    out = tmp_path / "out.docx"
    before_revs = rev_elements(doc)
    result = delete_citation(doc, out, paragraph_index=1, current_text="22")
    assert result["tracked"] is True and result["also_deleted"] == " "
    assert result["text_before"] == "Figure 2 legend 22." and result["text_after"] == "Figure 2 legend."
    xml = doc_xml(out)
    new = re.search(r'<w:del w:id="(\d+)" w:author="Claude" w:date="[^"]+">(.*?)</w:del>', xml, re.S)
    assert new and new.group(1) == result["revision_id"]
    inner = new.group(2)
    assert '<w:delText xml:space="preserve"> </w:delText>' in inner and "<w:delText>22</w:delText>" in inner
    assert "<w:delInstrText" in inner and "<w:instrText" not in inner and "<w:t>" not in inner
    assert inner.count('w:fldCharType="begin"') == 1 and inner.count('w:fldCharType="end"') == 1
    assert int(result["revision_id"]) > 102  # unique in the package
    assert '<w:t xml:space="preserve">Figure 2 legend</w:t>' in xml
    # every other revision is byte-identical
    assert [r for r in rev_elements(out) if f'w:id="{result["revision_id"]}"' not in r] == before_revs
    # K1 is still cited by cM and cK: nothing leaves the bibliography
    assert result["orphaned_items"] == [] and result["warnings"] == []
    # Reject (the MCP's own resolve) gives the live field back; Accept removes it
    rejected, accepted = tmp_path / "rej.docx", tmp_path / "acc.docx"
    resolve_revisions(out, rejected, mode="reject", ids=[result["revision_id"]])
    resolve_revisions(out, accepted, mode="accept", ids=[result["revision_id"]])
    assert instr_payloads(rejected) == instr_payloads(doc)
    assert para_texts(rejected)[0] == "Figure 2 legend 22."
    assert len(instr_payloads(accepted)) == len(instr_payloads(doc)) - 1
    assert "cX" not in "".join(instr_payloads(accepted))
    assert para_texts(accepted)[0] == "Figure 2 legend."


def test_orphaned_item_named_with_its_bibliography_entry(tmp_path):
    """A multi-item cluster where only one item is cited nowhere else: only that one is reported."""
    src = make_docx(tmp_path / "in.docx", t("See ") + field("cM", ["K1", "K3"], "(1,3)") + t("."),
                    t("Also ") + field("cK", ["K1"], "(1)") + t("."))
    result = delete_citation(src, tmp_path / "out.docx", paragraph_index=1, citation_id="cM")
    assert [(o["item_key"], o["bibliography_entry"]) for o in result["orphaned_items"]] == [("K3", "3. Title K3")]
    assert "leave the bibliography at the next refresh_citations after this deletion is accepted" in \
        result["warnings"][0] and "Title K3" in result["warnings"][0]
    assert result["text_after"] == "See."


def test_superscript_citation_without_space_keeps_neighbours(tmp_path):
    sup = "<w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr>"
    src = make_docx(tmp_path / "in.docx", t("accurate") + field("cS", ["K2"], "2", sup) + t(" results."))
    result = delete_citation(src, tmp_path / "out.docx", paragraph_index=1, citation_id="cS",
                             author="Editor Bot")
    assert result["also_deleted"] == "" and result["text_after"] == "accurate results."
    assert 'w:author="Editor Bot"' in doc_xml(tmp_path / "out.docx")


def test_space_kept_when_a_word_follows(tmp_path):
    src = make_docx(tmp_path / "in.docx", t("as shown ") + field("cW", ["K2"], "[2]") + t("above"))
    result = delete_citation(src, tmp_path / "out.docx", paragraph_index=1, citation_id="cW")
    assert result["also_deleted"] == "" and result["text_after"] == "as shown above"


def test_deleted_break_after_citation_is_not_the_next_character(tmp_path):
    """A tracked-deleted line break after the citation is not in the accepted view, so '.' follows it."""
    deleted_break = f'<w:del w:id="9" w:author="Sam" w:date="{DATE}"><w:r><w:br/></w:r></w:del>'
    src = make_docx(tmp_path / "in.docx", t("legend ") + field("cB", ["K2"], "22") + deleted_break + t("."))
    result = delete_citation(src, tmp_path / "out.docx", paragraph_index=1, citation_id="cB")
    assert result["also_deleted"] == " " and result["text_after"] == "legend."


def _simple(cid: str, keys: list[str], rendered: str) -> str:
    return f'<w:fldSimple w:instr="{_instr(cid, keys, rendered)}">{t(rendered)}</w:fldSimple>'


def test_simple_field_citation_is_refused_by_name(tmp_path):
    """Zotero's Word plugin writes complex fields; a w:fldSimple citation is named in the refusal, nothing written."""
    src = make_docx(tmp_path / "in.docx", t("See ") + _simple("cS", ["K1"], "(1)") + t("."), bibl=False)
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match=r"1 Zotero citation\(s\) stored as a simple field, w:fldSimple"):
        delete_citation(src, out, paragraph_index=1, current_text="(1)")
    assert not out.exists()


def test_item_still_cited_by_a_simple_field_is_not_orphaned(tmp_path):
    src = make_docx(tmp_path / "in.docx", t("See ") + field("cX", ["K1", "K2"], "(1,2)") + t("."),
                    t("Also ") + _simple("cS", ["K1"], "(1)") + t("."))
    result = delete_citation(src, tmp_path / "out.docx", paragraph_index=1, citation_id="cX")
    assert [item["item_key"] for item in result["orphaned_items"]] == ["K2"]


def test_untracked_delete_removes_the_field(doc, tmp_path):
    out = tmp_path / "out.docx"
    result = delete_citation(doc, out, paragraph_index=1, current_text="22", tracked=False)
    xml = doc_xml(out)
    assert result["tracked"] is False and result["revision_id"] is None
    assert "cX" not in xml and f'w:id="{int(103)}"' not in xml
    assert para_texts(out)[0] == "Figure 2 legend."
    assert rev_elements(out) == rev_elements(doc)


@pytest.mark.parametrize(("paragraphs", "kwargs", "message"), [
    ((t("x ") + f'<w:ins w:id="5" w:author="Sam" w:date="{DATE}">' + field("cI", ["K1"], "(1)") + "</w:ins>",),
     {"citation_id": "cI"}, "inside tracked revision"),
    ((t("x ") + field("cA", ["K1"], "(1)") + t(" y ") + field("cB", ["K2"], "(1)"),),
     {"current_text": "(1)"}, "2 Zotero citations in paragraph 1 match"),
    ((t("x ") + field("cA", ["K1"], "(1)"),), {"citation_id": "nope"}, "citations in that paragraph: cA"),
    ((t("x ") + field("cA", ["K1"], "(1)").replace('<w:r><w:t>', '<w:commentRangeStart w:id="7"/><w:r><w:t>'),),
     {"citation_id": "cA"}, "a comment marker lies inside citation cA"),
    ((t("x ") + field("cA", ["K1"], "(1)"),), {"citation_id": "cA", "current_text": "(1)"}, "exactly one"),
    ((t("x ") + field("cN", ["K1"], "(1)").replace(
        '<w:r><w:t>', '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>'),), {"citation_id": "cN"}, "nested field"),
    ((t("x ") + field("cP", ["K1"], "(1)").replace(
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>', '</w:p><w:p><w:r><w:fldChar w:fldCharType="end"/></w:r>'),),
     {"citation_id": "cP"}, "spans paragraphs"),
    ((t("Legend ") + field("cR", ["K1"], "22").replace(
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>', '<w:r><w:fldChar w:fldCharType="end"/><w:t>.</w:t></w:r>'),),
     {"citation_id": "cR"}, "shares a run with <w:t> outside the field"),
    ((field("cL", ["K1"], "22").replace(
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
        '<w:r><w:t xml:space="preserve">Legend </w:t><w:fldChar w:fldCharType="begin"/></w:r>', 1) + t("."),),
     {"citation_id": "cL"}, "shares a run with <w:t> outside the field"),
], ids=["in-insertion", "ambiguous", "missing", "partial-comment", "both-selectors", "nested", "two-paragraphs",
        "text-after-end-in-run", "text-before-begin-in-run"])
def test_refusals_name_the_reason(tmp_path, paragraphs, kwargs, message):
    src = make_docx(tmp_path / "in.docx", *paragraphs, bibl=False)
    out = tmp_path / "out.docx"
    with pytest.raises(RefusedUnsafe, match=re.escape(message)):
        delete_citation(src, out, paragraph_index=1, **kwargs)
    assert not out.exists()


def test_comment_wholly_on_the_citation_keeps_its_markers_outside(tmp_path):
    src = make_docx(tmp_path / "in.docx", t("Legend ") + '<w:commentRangeStart w:id="7"/>'
                    + field("cC", ["K2"], "(2)") + '<w:commentRangeEnd w:id="7"/>'
                    + '<w:r><w:commentReference w:id="7"/></w:r>' + t("."))
    out = tmp_path / "out.docx"
    result = delete_citation(src, out, paragraph_index=1, citation_id="cC")
    xml = doc_xml(out)
    deletion = re.search(r"<w:del w:id[^>]*>.*?</w:del>", xml, re.S).group(0)
    assert "commentRange" not in deletion and "commentReference" not in deletion
    assert xml.index('<w:commentRangeStart w:id="7"/>') < xml.index(deletion) < xml.index(
        '<w:commentRangeEnd w:id="7"/>')
    assert result["text_after"] == "Legend ."
    assert any("accepting the deletion leaves comment(s) ['7'] with no text" in w for w in result["warnings"])


def test_untracked_delete_warns_about_a_comment_left_without_text(tmp_path):
    src = make_docx(tmp_path / "in.docx", t("Legend ") + '<w:commentRangeStart w:id="7"/>'
                    + field("cC", ["K2"], "(2)") + '<w:commentRangeEnd w:id="7"/>'
                    + '<w:r><w:commentReference w:id="7"/></w:r>' + t("."))
    result = delete_citation(src, tmp_path / "out.docx", paragraph_index=1, citation_id="cC", tracked=False)
    assert any("the deletion leaves comment(s) ['7'] with no text" in w for w in result["warnings"])


windows_only = pytest.mark.skipif(__import__("os").name != "nt", reason="undo store writes need Windows Python")


@windows_only
def test_edit_document_delete_citation_in_place_with_undo(doc):
    """Through edit_document on a document with other revisions: one undoable batch, details as the op."""
    from research_mcp.audit import file_sha256
    from research_mcp.edit_batch import edit_document
    from research_mcp.edit_models import EditDocumentRequest
    from research_mcp.publication import undo_document

    before = file_sha256(doc)
    result = edit_document(EditDocumentRequest(source=str(doc), source_sha256=before, output=str(doc), operations=[
        {"op_id": "d", "type": "delete_citation", "paragraph_index": 3, "citation_id": "cM"}]))
    assert result.status == "complete", result
    details = result.operations[0].details
    assert details["citation_id"] == "cM" and details["revision_id"]
    assert [o["item_key"] for o in details["orphaned_items"]] == ["K3"]
    assert details["text_before"] == "Both (1,3) and (1)x" and details["text_after"] == "Both and (1)x"
    assert para_texts(doc)[2] == "Both and (1)x"
    assert len(instr_payloads(doc)) == 3  # cX, cB, cK live; cM's field code is now delInstrText
    undo_document(str(doc))
    assert file_sha256(doc) == before


def test_server_envelope_maps_delete_citation():
    from research_mcp.server import MCPEditOperation, _strict_operation_payload

    op = _strict_operation_payload(MCPEditOperation(op_id="d", type="delete_citation", paragraph_index=4,
                                                    current_text="22"))
    assert (op.type, op.paragraph_index, op.current_text, op.citation_id, op.author, op.tracked) == \
        ("delete_citation", 4, "22", None, "Claude", True)
    untracked = _strict_operation_payload(MCPEditOperation(op_id="d", type="delete_citation", citation_id="cX",
                                                           tracked=False, author="Editor Bot"))
    assert (untracked.citation_id, untracked.tracked, untracked.author) == ("cX", False, "Editor Bot")


def test_empty_item_keys_points_at_delete_citation(tmp_path):
    from research_mcp.cluster_ops import edit_cluster_to_version

    src = make_docx(tmp_path / "in.docx", LEGEND)
    with pytest.raises(RefusedUnsafe, match="delete_citation"):
        edit_cluster_to_version(src, src, source_sha256="0" * 64, paragraph_index=1, item_keys=[],
                                current_text="22")
