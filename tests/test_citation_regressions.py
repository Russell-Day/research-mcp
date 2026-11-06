from types import SimpleNamespace

from research_mcp.citations import (
    BIBLIOGRAPHY_CODE,
    BIBLIOGRAPHY_PLACEHOLDER,
    _reset_bibliography_field,
    renew_zotero_session,
)
from research_mcp.word_worker import (
    _bibliographies_complete,
    _mark_refresh_pending,
    _refresh_completed,
)


def test_cloned_bibliography_is_reset_to_neutral_placeholder():
    field = SimpleNamespace(
        Code=SimpleNamespace(Text="ADDIN ZOTERO_BIBL donor code"),
        Result=SimpleNamespace(Text="Donor rendered entry"),
    )
    _reset_bibliography_field(field)
    assert field.Code.Text == BIBLIOGRAPHY_CODE
    assert field.Result.Text == BIBLIOGRAPHY_PLACEHOLDER


def test_empty_bibliography_is_complete_but_placeholder_is_not():
    assert _bibliographies_complete([])
    assert _bibliographies_complete([""])
    assert not _bibliographies_complete([BIBLIOGRAPHY_PLACEHOLDER])


class FakeProperties:
    def __init__(self, values):
        self._items = [SimpleNamespace(Name=name, Value=value) for name, value in values]
        self.Count = len(self._items)

    def __call__(self, index):
        return self._items[index - 1]


def test_new_document_version_gets_unique_zotero_session():
    properties = FakeProperties([
        ("ZOTERO_PREF_2", "/data>"),
        ("ZOTERO_PREF_1", '<data><session id="OLDSESS1"/><prefs>ok</prefs><'),
    ])
    document = SimpleNamespace(CustomDocumentProperties=properties)
    old_id, new_id = renew_zotero_session(document)
    combined = properties(2).Value + properties(1).Value
    assert old_id == "OLDSESS1"
    assert new_id != old_id
    assert f'<session id="{new_id}"/>' in combined


def test_refresh_completion_requires_zotero_to_replace_marker():
    properties = FakeProperties([
        ("ZOTERO_PREF_2", '<style id="x"/></data>'),
        ("ZOTERO_PREF_1", '<data data-version="3" zotero-version="9.0.6"><session id="ABC"/>'),
    ])
    document = SimpleNamespace(CustomDocumentProperties=properties)
    marker = _mark_refresh_pending(document)
    assert not _refresh_completed(document, marker)

    combined = properties(1).Value + properties(2).Value
    rewritten = combined.replace(marker, "9.0.6")
    split = len(properties(1).Value)
    properties(1).Value = rewritten[:split]
    properties(2).Value = rewritten[split:]
    assert _refresh_completed(document, marker)
