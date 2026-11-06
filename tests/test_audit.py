from __future__ import annotations

import xml.etree.ElementTree as ET
import zipfile

from research_mcp.audit import audit_docx, cell_text_digest, document_structure_errors
from research_mcp.ooxml import NS, parse_table


def test_audit_reads_merge_and_bookmark_reference(minimal_docx):
    audit = audit_docx(minimal_docx)
    assert audit.table_count == 1
    assert audit.tables[0].rows == 3
    assert audit.tables[0].columns == 2
    assert audit.tables[0].vertical_merges == ((1, 3, 1, "LoRA"),)
    assert audit.bookmark_names == ("KeepMe",)
    assert audit.bookmark_references == {"KeepMe": ("hyperlink",)}
    assert audit.media_count == 1
    assert audit.tables[0].cell_hashes[0] == audit.tables[0].cell_hashes[2]


def test_grid_span_expands_to_logical_columns():
    table = ET.fromstring(
        """<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid>
        <w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"""
    )
    logical = parse_table(table)
    assert logical.cells == (("A", "A", "B"),)
    assert logical.horizontal_spans == ((1, 1, 2),)



def test_agent_dict_drops_per_cell_hashes_for_one_digest_per_table(minimal_docx):
    """The MCP boundary must not ship verifier-only data to the model."""

    audit = audit_docx(minimal_docx)
    full = audit.to_dict()
    agent = audit.to_agent_dict()

    assert "cell_hashes" in full["tables"][0]
    assert "cell_hashes" not in agent["tables"][0]
    assert agent["tables"][0]["cell_text_sha256"] == cell_text_digest(
        audit.tables[0].cell_hashes
    )

    # Everything else is preserved, including bookmark_names: an agent needs
    # exact names to build a safe clean_bookmarks call, and validate_release
    # refuses any name absent from the audit.
    assert agent["bookmark_names"] == ("KeepMe",)
    assert set(full) == set(agent)
    assert set(full["tables"][0]) - set(agent["tables"][0]) == {"cell_hashes"}
    assert set(agent["tables"][0]) - set(full["tables"][0]) == {"cell_text_sha256"}


def test_cell_text_digest_detects_any_cell_change_and_any_reordering():
    """The digest has to carry the integrity claim the hashes used to carry."""

    base = ("A", "B", "C")
    assert cell_text_digest(base) == cell_text_digest(("A", "B", "C"))
    assert cell_text_digest(base) != cell_text_digest(("A", "B", "D"))
    assert cell_text_digest(base) != cell_text_digest(("A", "C", "B"))
    # A separator is required, or ("AB","C") and ("A","BC") would collide.
    assert cell_text_digest(("AB", "C")) != cell_text_digest(("A", "BC"))
    assert cell_text_digest(()) == cell_text_digest([])


def test_core_verification_still_sees_the_full_audit(minimal_docx):
    """Dropping hashes at the boundary must not weaken tables.py or verifiers."""

    audit = audit_docx(minimal_docx)
    assert audit.tables[0].cell_hashes
    assert audit.to_dict()["tables"][0]["cell_hashes"]


def test_structure_preflight_reports_undeclared_ignorable_and_unclosed_field(tmp_path):
    path = tmp_path / "broken.docx"
    document = """<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14">
      <w:body><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p><w:sectPr/></w:body></w:document>"""
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", document)

    errors = document_structure_errors(path)
    assert any("undeclared mc:Ignorable prefix: w14" in error for error in errors)
    assert any("complex field(s) are unclosed" in error for error in errors)
    audit = audit_docx(path)
    assert audit.citation_readiness["ready"] is False
    assert audit.citation_readiness["structural_errors"] == list(errors)


def test_structure_preflight_handles_xml_declaration_and_arbitrary_mc_alias(tmp_path):
    path = tmp_path / "namespaces.docx"
    document = """<?xml version="1.0" encoding="UTF-8"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:c="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:good="urn:example:good" c:Ignorable="w14 good">
      <w:body><w:p/><w:sectPr/></w:body></w:document>"""
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", document)

    errors = document_structure_errors(path)
    assert errors == ("word/document.xml: undeclared mc:Ignorable prefix: w14",)


def test_structure_preflight_binds_qname_values_and_nested_scope(tmp_path):
    path = tmp_path / "qnames.docx"
    document = """<?xml version="1.0" encoding="UTF-8"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:c="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="urn:w14" c:Ignorable="w14">
      <w:body><c:Choice Requires="w14"><c:Fallback xmlns:ok="urn:ok" c:PreserveElements="ok:element" c:ProcessContent="bad:element"/></c:Choice><w:sectPr/></w:body></w:document>"""
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", document)

    errors = document_structure_errors(path)
    assert "word/document.xml: undeclared QName prefix in Requires: w14" not in errors
    assert "word/document.xml: undeclared QName prefix in PreserveElements: ok" not in errors
    assert any("ProcessContent prefix: bad" in error for error in errors)


def test_structure_preflight_reports_missing_document_and_bad_zip(tmp_path):
    missing = tmp_path / "missing.docx"
    with zipfile.ZipFile(missing, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
    assert document_structure_errors(missing) == ("package: missing word/document.xml",)

    bad = tmp_path / "bad.docx"
    bad.write_bytes(b"not a zip")
    assert document_structure_errors(bad)[0].startswith("package: invalid ZIP:")


def test_media_relationships_are_namespace_and_order_insensitive(tmp_path):
    path = tmp_path / "media.docx"
    document = """<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:ns="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <w:body><w:p><w:r><w:drawing ns:embed="rId1"/></w:r></w:p><w:sectPr/></w:body></w:document>"""
    rels = """<ns:Relationships xmlns:ns="http://schemas.openxmlformats.org/package/2006/relationships">
      <ns:Relationship Target="media/image1.png" Type="x" Id="rId1"/></ns:Relationships>"""
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", document)
        archive.writestr("word/_rels/document.xml.rels", rels)
        archive.writestr("word/media/image1.png", b"fixture")

    audit = audit_docx(path)
    assert audit.media_count == 1
    assert audit.live_media_count == 1
    assert audit.dangling_media_targets == ()
