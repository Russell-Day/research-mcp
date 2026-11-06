from pathlib import Path

from research_mcp.audit import audit_docx
from research_mcp.ooxml import zotero_package_fields


DATA = Path(__file__).parent / "data"


def test_project_owned_item_donor_is_exact():
    audit = audit_docx(DATA / "zotero_item_donor.docx")
    assert audit.sha256 == "7C454923542535A3D7A18EE8401453867E6CCA363AAF35DE180174A2CE670301"
    assert audit.zotero_item_fields == 1
    assert audit.zotero_bibliography_fields == 0
    assert audit.zotero_pref_properties == 2


def test_project_owned_bibliography_donor_is_exact():
    audit = audit_docx(DATA / "zotero_bibliography_donor.docx")
    assert audit.sha256 == "07B356B1F582806A5EE174D7FB4ACAB9EACF4DD6D97AC72E221F241EC6759684"
    assert audit.zotero_item_fields == 0
    assert audit.zotero_bibliography_fields == 1
    assert audit.zotero_pref_properties == 2


def test_project_owned_integration_source_is_field_free():
    audit = audit_docx(DATA / "zotero_integration_source.docx")
    assert audit.sha256 == "35B862E399646EA708C5BF7F8664F76380DFD843A244970310FEDD6D30958B03"
    assert audit.zotero_item_fields == 0
    assert audit.zotero_bibliography_fields == 0
    assert audit.zotero_pref_properties == 2


def test_closed_package_field_reader_handles_project_donors():
    item = zotero_package_fields(DATA / "zotero_item_donor.docx")
    bibliography = zotero_package_fields(DATA / "zotero_bibliography_donor.docx")
    assert len(item.item_results) == 1
    assert item.item_results[0][1] == "(1)"
    assert item.item_keys == (("6W98BDK7",),)
    assert item.bibliography_results == ()
    assert bibliography.item_results == ()
    assert bibliography.item_keys == ()
    assert bibliography.bibliography_results == ("{Bibliography}",)
