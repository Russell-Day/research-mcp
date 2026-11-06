from __future__ import annotations

import io
import json
from urllib.error import HTTPError

import pytest

from research_mcp.citations import PLACEHOLDER, build_payload
from research_mcp.errors import RefusedUnsafe, ZoteroItemNotFound, ZoteroUnavailable
from research_mcp.zotero import (
    ZoteroClient,
    add_item_by_doi,
    find_project_collection,
    find_project_collection_details,
    normalize_doi,
    write_project_collection,
)


class Response(io.BytesIO):
    status = 200

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()


class FakeWriteClient:
    """Stand-in for pyzotero's write client -- no real network, no real writes."""

    def __init__(self, *, item_key: str = "NEWKEY1", collection_key: str = "COLKEY01"):
        self.next_item_key = item_key
        self.next_collection_key = collection_key
        self.created_items: list[dict] = []
        self.created_collections: list[dict] = []
        self.added_to_collection: list[tuple[str, dict]] = []
        self.attachments: list[tuple[tuple[str, ...], str]] = []

    def item_template(self, item_type):
        return {
            "itemType": item_type, "title": "", "creators": [], "date": "",
            "DOI": "", "url": "", "volume": "", "issue": "", "pages": "",
            "publisher": "", "ISSN": "", "publicationTitle": "", "abstractNote": "",
        }

    def create_items(self, items):
        self.created_items.extend(items)
        return {"success": {"0": self.next_item_key}}

    def create_collections(self, collections):
        self.created_collections.extend(collections)
        return {"success": {"0": self.next_collection_key}}

    def addto_collection(self, collection_key, item):
        self.added_to_collection.append((collection_key, item))

    def item(self, key):
        return {"key": key, "data": {"key": key}}

    def attachment_simple(self, paths, item_key):
        self.attachments.append((tuple(paths), item_key))


def test_build_payload_requires_stable_item_ids():
    with pytest.raises(RefusedUnsafe):
        build_payload([])
    with pytest.raises(RefusedUnsafe):
        build_payload([{"title": "Missing id"}])


def test_build_payload_is_unrendered_and_keeps_item_data():
    item = {"id": "http://zotero.org/users/local/abc/items/KEY123", "title": "Paper"}
    payload = build_payload([item], citation_id="CITE1234")
    assert payload["citationID"] == "CITE1234"
    assert payload["properties"]["formattedCitation"] == PLACEHOLDER
    assert payload["citationItems"][0]["itemData"] == item


def test_zotero_client_search_and_csl(monkeypatch):
    calls = []

    def fake_urlopen(request, timeout):
        calls.append((request.full_url, dict(request.header_items()), timeout))
        if "format=csljson" in request.full_url:
            payload = [{"id": "http://zotero.org/users/local/a/items/K1", "title": "CSL"}]
        elif request.full_url.rstrip("/").endswith("/items/K1"):
            # csl() checks the trash flag before returning citable data.
            payload = {"key": "K1", "data": {"key": "K1", "title": "One"}}
        else:
            payload = [
                {"key": "K1", "data": {"title": "One", "date": "2024", "creators": [{"lastName": "Doe"}]}},
                {"key": "A1", "data": {"itemType": "attachment", "title": "PDF"}},
            ]
        return Response(json.dumps(payload).encode())

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)
    client = ZoteroClient(timeout=3)
    assert client.csl("K1")["title"] == "CSL"
    assert client.search("one")[0].authors == "Doe"
    assert all(headers["Zotero-allowed-request"] == "1" for _, headers, _ in calls)


def test_zotero_client_wraps_transport_errors(monkeypatch):
    def fail(*args, **kwargs):
        raise OSError("offline")

    monkeypatch.setattr("research_mcp.zotero.urlopen", fail)
    with pytest.raises(ZoteroUnavailable, match="offline"):
        ZoteroClient().csl("K1")


def test_csl_refuses_a_trashed_item(monkeypatch):
    """Search hides trashed items but a direct key fetch does not.

    Without this guard a stale key produces a live Word citation pointing at an
    item that disappears when the user empties the trash.
    """

    def fake_urlopen(request, timeout):
        if "format=csljson" in request.full_url:
            payload = [{"id": "http://zotero.org/users/local/a/items/GONE", "title": "Trashed"}]
        else:
            payload = {"key": "GONE", "data": {"key": "GONE", "deleted": True, "title": "Trashed"}}
        return Response(json.dumps(payload).encode())

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)
    with pytest.raises(RefusedUnsafe, match="in the trash"):
        ZoteroClient(timeout=3).csl("GONE")


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("10.1038/nature14539", "10.1038/nature14539"),
        ("https://doi.org/10.1038/NATURE14539", "10.1038/nature14539"),
        ("doi: 10.1038/Nature14539 ", "10.1038/nature14539"),
    ],
)
def test_normalize_doi_accepts_common_forms(raw, expected):
    assert normalize_doi(raw) == expected


@pytest.mark.parametrize("raw", ["", "nature14539", "10.1038", "not a doi"])
def test_normalize_doi_refuses_non_dois(raw):
    with pytest.raises(RefusedUnsafe):
        normalize_doi(raw)


def test_find_by_doi_scans_items_and_skips_trash(monkeypatch):
    """The search endpoint does not index DOI, so lookup must scan top-level items."""

    page = [
        {"key": "LIVE", "data": {"key": "LIVE", "DOI": "10.1073/PNAS.1611835114"}},
        {"key": "TRASH", "data": {"key": "TRASH", "DOI": "10.1038/nature14539", "deleted": True}},
        {"key": "NODOI", "data": {"key": "NODOI"}},
    ]

    def fake_urlopen(request, timeout):
        body = page if "start=0" in request.full_url else []
        return Response(json.dumps(body).encode())

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)
    client = ZoteroClient(timeout=3)
    assert client.find_by_doi("https://doi.org/10.1073/pnas.1611835114") == "LIVE"
    assert client.find_by_doi("10.1038/nature14539") is None


def test_find_by_doi_refuses_duplicate_matches(monkeypatch):
    page = [
        {"key": "A", "data": {"key": "A", "DOI": "10.1/x"}},
        {"key": "B", "data": {"key": "B", "DOI": "10.1/X"}},
    ]

    def fake_urlopen(request, timeout):
        return Response(json.dumps(page if "start=0" in request.full_url else []).encode())

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)
    with pytest.raises(RefusedUnsafe, match="multiple Zotero items"):
        ZoteroClient(timeout=3).find_by_doi("10.1/x")


# -- collection resolution ----------------------------------------------------

def test_resolve_collection_accepts_exact_key(monkeypatch):
    def fail_urlopen(request, timeout):
        raise AssertionError("should not call the API for a key-shaped input")

    monkeypatch.setattr("research_mcp.zotero.urlopen", fail_urlopen)
    assert ZoteroClient(timeout=3).resolve_collection("ABCD1234") == "ABCD1234"


def test_resolve_collection_lists_available_names_when_missing(monkeypatch):
    def fake_urlopen(request, timeout):
        page = [{"key": "AAAAAAAA", "data": {"key": "AAAAAAAA", "name": "Alpha", "parentCollection": False}}]
        return Response(json.dumps(page if "start=0" in request.full_url else []).encode())

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)
    with pytest.raises(RefusedUnsafe, match=r"available: \['Alpha'\]"):
        ZoteroClient(timeout=3).resolve_collection("Missing")


def test_resolve_collection_suggests_close_match_with_parent(monkeypatch):
    def fake_urlopen(request, timeout):
        page = [
            {"key": "PPPPPPPP", "data": {"key": "PPPPPPPP", "name": "Projects", "parentCollection": False}},
            {
                "key": "AAAAAAAA",
                "data": {"key": "AAAAAAAA", "name": "Dissertation", "parentCollection": "PPPPPPPP"},
            },
        ]
        return Response(json.dumps(page if "start=0" in request.full_url else []).encode())

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)
    with pytest.raises(RefusedUnsafe, match=r"did you mean 'Dissertation' \(under 'Projects'\)"):
        ZoteroClient(timeout=3).resolve_collection("Disertation")


def test_resolve_collection_no_hint_when_no_close_match(monkeypatch):
    def fake_urlopen(request, timeout):
        page = [{"key": "AAAAAAAA", "data": {"key": "AAAAAAAA", "name": "Alpha", "parentCollection": False}}]
        return Response(json.dumps(page if "start=0" in request.full_url else []).encode())

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)
    with pytest.raises(RefusedUnsafe) as excinfo:
        ZoteroClient(timeout=3).resolve_collection("Completely Unrelated Name")
    assert "did you mean" not in str(excinfo.value)


def test_resolve_collection_raises_on_ambiguous_name(monkeypatch):
    def fake_urlopen(request, timeout):
        page = [
            {"key": "AAAAAAAA", "data": {"key": "AAAAAAAA", "name": "Notes", "parentCollection": False}},
            {"key": "BBBBBBBB", "data": {"key": "BBBBBBBB", "name": "Notes", "parentCollection": "AAAAAAAA"}},
        ]
        return Response(json.dumps(page if "start=0" in request.full_url else []).encode())

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)
    with pytest.raises(RefusedUnsafe, match="ambiguous"):
        ZoteroClient(timeout=3).resolve_collection("Notes")


def test_search_with_collection_key_hits_scoped_endpoint(monkeypatch):
    seen_urls = []

    def fake_urlopen(request, timeout):
        seen_urls.append(request.full_url)
        return Response(json.dumps([]).encode())

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)
    ZoteroClient(timeout=3).search("term", collection_key="COLKEY1")
    assert seen_urls[0].startswith("http://127.0.0.1:23119/api/users/0/collections/COLKEY1/items?")


# -- .zotero-collection marker file -------------------------------------------

def test_find_project_collection_reads_marker_in_doc_folder(tmp_path):
    (tmp_path / ".zotero-collection").write_text("Thesis\n", encoding="utf-8")
    doc = tmp_path / "draft.docx"
    doc.write_bytes(b"")
    assert find_project_collection(str(doc)) == "Thesis"


def test_find_project_collection_walks_up_to_parent(tmp_path):
    (tmp_path / ".zotero-collection").write_text("Parent Project\n", encoding="utf-8")
    sub = tmp_path / "drafts" / "v2"
    sub.mkdir(parents=True)
    doc = sub / "draft.docx"
    doc.write_bytes(b"")
    assert find_project_collection(str(doc)) == "Parent Project"


def test_find_project_collection_absent_returns_none(tmp_path):
    doc = tmp_path / "draft.docx"
    doc.write_bytes(b"")
    assert find_project_collection(str(doc)) is None


def test_find_project_collection_details_returns_inherited_value_and_path(tmp_path):
    marker = tmp_path / ".zotero-collection"
    marker.write_text("Parent Project\n", encoding="utf-8")
    sub = tmp_path / "drafts"
    sub.mkdir()
    doc = sub / "draft.docx"
    doc.write_bytes(b"")

    assert find_project_collection_details(str(doc)) == ("Parent Project", marker.resolve())


def test_find_project_collection_skips_blank_marker_to_reach_grandparent(tmp_path):
    grandparent_marker = tmp_path / ".zotero-collection"
    grandparent_marker.write_text("Grandparent Project\n", encoding="utf-8")
    parent = tmp_path / "drafts"
    parent.mkdir()
    (parent / ".zotero-collection").write_text("\n", encoding="utf-8")
    sub = parent / "v2"
    sub.mkdir()
    doc = sub / "draft.docx"
    doc.write_bytes(b"")

    assert find_project_collection(str(doc)) == "Grandparent Project"
    assert find_project_collection_details(str(doc)) == (
        "Grandparent Project",
        grandparent_marker.resolve(),
    )


def test_write_project_collection_writes_document_parent_and_reports_no_previous(tmp_path):
    folder = tmp_path / "drafts"
    folder.mkdir()
    doc = folder / "draft.docx"
    doc.write_bytes(b"")

    result = write_project_collection(str(doc), "Thesis")

    assert result == {
        "marker_path": str((folder / ".zotero-collection").resolve()),
        "previous_value": None,
        "replaced": False,
    }
    assert (folder / ".zotero-collection").read_text(encoding="utf-8") == "Thesis\n"


def test_write_project_collection_refuses_unexpected_replacement(tmp_path):
    marker = tmp_path / ".zotero-collection"
    marker.write_text("Existing\n", encoding="utf-8")

    with pytest.raises(RefusedUnsafe, match="current value is 'Existing'"):
        write_project_collection(str(tmp_path), "New", replace_existing=False)

    assert marker.read_text(encoding="utf-8") == "Existing\n"


def test_write_project_collection_replaces_only_when_explicit(tmp_path):
    marker = tmp_path / ".zotero-collection"
    marker.write_text("Existing\n", encoding="utf-8")

    result = write_project_collection(str(tmp_path), "New", replace_existing=True)

    assert result["previous_value"] == "Existing"
    assert result["replaced"] is True
    assert marker.read_text(encoding="utf-8") == "New\n"


# -- add_item_by_doi -----------------------------------------------------------

def test_add_item_by_doi_dedups_existing_without_write_client(monkeypatch):
    def fake_urlopen(request, timeout):
        url = request.full_url
        if "/items/top" in url:
            page = [{"key": "K1", "data": {"key": "K1", "DOI": "10.1000/existing"}}]
            return Response(json.dumps(page if "start=0" in url else []).encode())
        if url.rstrip("/").endswith("/items/K1"):
            data = {
                "key": "K1", "title": "Existing Paper",
                "creators": [{"creatorType": "author", "lastName": "Doe"}],
                "abstractNote": "An abstract.",
            }
            return Response(json.dumps({"key": "K1", "data": data}).encode())
        raise AssertionError(f"unexpected local API call: {url}")

    def fail_write_client():
        raise AssertionError("write client should not be constructed without a collection")

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)
    monkeypatch.setattr("research_mcp.zotero._write_client_from_env", fail_write_client)

    result = add_item_by_doi("10.1000/existing")
    assert result == {
        "success": True, "created": False, "item_key": "K1", "doi": "10.1000/existing",
        "title": "Existing Paper", "first_author": "Doe", "abstract": "An abstract.",
        "pdf_attached": False, "pdf_status": "skipped (existing item)",
        "visible_locally": True, "collection": None, "collection_marker_path": None,
    }


def test_add_item_by_doi_reports_inherited_marker_path(monkeypatch, tmp_path):
    marker = tmp_path / ".zotero-collection"
    marker.write_text("Thesis\n", encoding="utf-8")
    sub = tmp_path / "drafts"
    sub.mkdir()
    doc = sub / "draft.docx"
    doc.write_bytes(b"")
    fake_write = FakeWriteClient()

    def fake_urlopen(request, timeout):
        url = request.full_url
        if "/items/top" in url:
            page = [{"key": "K1", "data": {"key": "K1", "DOI": "10.1000/existing"}}]
            return Response(json.dumps(page if "start=0" in url else []).encode())
        if "/collections?limit=100" in url:
            page = [{
                "key": "THESIS01",
                "data": {"key": "THESIS01", "name": "Thesis", "parentCollection": False},
            }]
            return Response(json.dumps(page if "start=0" in url else []).encode())
        if url.rstrip("/").endswith("/items/K1"):
            data = {
                "key": "K1", "title": "Existing Paper",
                "creators": [{"creatorType": "author", "lastName": "Doe"}],
                "abstractNote": "An abstract.",
            }
            return Response(json.dumps({"key": "K1", "data": data}).encode())
        raise AssertionError(f"unexpected local API call: {url}")

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)
    monkeypatch.setattr("research_mcp.zotero._write_client_from_env", lambda: fake_write)

    result = add_item_by_doi("10.1000/existing", path=str(doc))

    assert result["collection"] == "Thesis"
    assert result["collection_marker_path"] == str(marker.resolve())
    assert fake_write.added_to_collection[0][0] == "THESIS01"


def test_add_item_by_doi_missing_collection_raises_without_create(monkeypatch):
    def fake_urlopen(request, timeout):
        page = [{"key": "AAAAAAAA", "data": {"key": "AAAAAAAA", "name": "Other", "parentCollection": False}}]
        return Response(json.dumps(page if "start=0" in request.full_url else []).encode())

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)
    with pytest.raises(RefusedUnsafe, match="no Zotero collection named 'Missing'"):
        add_item_by_doi("10.1/x", collection="Missing")


def test_add_item_by_doi_creates_missing_collection_when_requested(monkeypatch):
    monkeypatch.delenv("ZOTERO_MCP_CONTACT_EMAIL", raising=False)
    fake_write = FakeWriteClient()
    monkeypatch.setattr("research_mcp.zotero._write_client_from_env", lambda: fake_write)
    monkeypatch.setattr(
        "research_mcp.zotero._fetch_crossref_work",
        lambda doi: {"type": "report", "title": ["Report Title"]},
    )

    def fake_urlopen(request, timeout):
        url = request.full_url
        if "/items/top" in url:
            return Response(json.dumps([]).encode())
        if "/collections?limit" in url:
            return Response(json.dumps([]).encode())
        if url.rstrip("/").endswith(f"/items/{fake_write.next_item_key}"):
            return Response(json.dumps({"key": fake_write.next_item_key, "data": {}}).encode())
        raise AssertionError(f"unexpected local API call: {url}")

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)

    result = add_item_by_doi("10.1000/new2", collection="New Folder", create_collection=True)
    assert result["collection"] == "New Folder"
    assert fake_write.created_collections == [{"name": "New Folder"}]
    assert fake_write.created_items[0]["collections"] == [fake_write.next_collection_key]


def test_add_item_by_doi_refuses_near_duplicate_even_with_create_collection(monkeypatch):
    def fail_write_client():
        raise AssertionError("write client should not be constructed on refusal")

    def fake_urlopen(request, timeout):
        page = [{"key": "AAAAAAAA", "data": {"key": "AAAAAAAA", "name": "Dissertation", "parentCollection": False}}]
        return Response(json.dumps(page if "start=0" in request.full_url else []).encode())

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)
    monkeypatch.setattr("research_mcp.zotero._write_client_from_env", fail_write_client)

    with pytest.raises(RefusedUnsafe, match=r"close to 'Dissertation' \(top-level\)"):
        add_item_by_doi("10.1/x", collection="Disertation", create_collection=True)


def test_add_item_by_doi_maps_crossref_fields_and_files_into_collection(monkeypatch):
    monkeypatch.delenv("ZOTERO_MCP_CONTACT_EMAIL", raising=False)
    fake_write = FakeWriteClient()
    monkeypatch.setattr("research_mcp.zotero._write_client_from_env", lambda: fake_write)
    monkeypatch.setattr(
        "research_mcp.zotero._fetch_crossref_work",
        lambda doi: {
            "type": "journal-article",
            "title": ["A Great Paper"],
            "author": [{"family": "Doe", "given": "Jane"}],
            "published": {"date-parts": [[2024, 3]]},
            "URL": "https://example.org/paper",
            "container-title": ["Journal of Examples"],
            "abstract": "<jats:p>An <b>abstract</b>.</jats:p>",
        },
    )

    def fake_urlopen(request, timeout):
        url = request.full_url
        if "/items/top" in url:
            return Response(json.dumps([]).encode())
        if "/collections?limit" in url:
            page = [{"key": "COLKEY1", "data": {"key": "COLKEY1", "name": "Thesis", "parentCollection": False}}]
            return Response(json.dumps(page if "start=0" in url else []).encode())
        if url.rstrip("/").endswith(f"/items/{fake_write.next_item_key}"):
            return Response(json.dumps({"key": fake_write.next_item_key, "data": {}}).encode())
        raise AssertionError(f"unexpected local API call: {url}")

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)

    result = add_item_by_doi("10.1000/new", collection="Thesis")
    assert result["created"] is True
    assert result["item_key"] == fake_write.next_item_key
    assert result["title"] == "A Great Paper"
    assert result["first_author"] == "Doe"
    assert result["abstract"] == "An abstract."
    assert result["pdf_attached"] is False
    assert result["pdf_status"] == "no contact email configured"
    assert result["collection"] == "Thesis"
    assert fake_write.created_items[0]["collections"] == ["COLKEY1"]
    assert fake_write.created_items[0]["date"] == "2024-3"


def test_add_item_by_doi_crossref_404_raises(monkeypatch):
    def fake_urlopen(request, timeout):
        url = request.full_url
        if url.startswith("http://127.0.0.1"):
            if "/items/top" in url:
                return Response(json.dumps([]).encode())
            raise AssertionError(f"unexpected local API call: {url}")
        if url.startswith("https://api.crossref.org"):
            raise HTTPError(url, 404, "Not Found", None, None)
        raise AssertionError(f"unexpected url: {url}")

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)
    monkeypatch.setattr("research_mcp.zotero._write_client_from_env", lambda: FakeWriteClient())
    with pytest.raises(ZoteroItemNotFound):
        add_item_by_doi("10.9999/missing")


def test_write_client_from_env_reports_missing_vars(monkeypatch):
    monkeypatch.delenv("ZOTERO_API_KEY", raising=False)
    monkeypatch.delenv("ZOTERO_LIBRARY_ID", raising=False)
    from research_mcp.zotero import _write_client_from_env

    with pytest.raises(ZoteroUnavailable, match="ZOTERO_LIBRARY_ID, ZOTERO_API_KEY"):
        _write_client_from_env()


def test_add_item_by_doi_attaches_oa_pdf_when_email_configured(monkeypatch):
    monkeypatch.setenv("ZOTERO_MCP_CONTACT_EMAIL", "test@example.org")
    fake_write = FakeWriteClient()
    monkeypatch.setattr("research_mcp.zotero._write_client_from_env", lambda: fake_write)
    monkeypatch.setattr(
        "research_mcp.zotero._fetch_crossref_work",
        lambda doi: {"type": "journal-article", "title": ["T"]},
    )
    monkeypatch.setattr(
        "research_mcp.zotero._find_unpaywall_pdf_url", lambda doi: "https://example.org/paper.pdf"
    )

    def fake_urlopen(request, timeout):
        url = request.full_url
        if url.startswith("http://127.0.0.1"):
            if "/items/top" in url:
                return Response(json.dumps([]).encode())
            if url.rstrip("/").endswith(f"/items/{fake_write.next_item_key}"):
                return Response(json.dumps({"key": fake_write.next_item_key, "data": {}}).encode())
            raise AssertionError(f"unexpected local API call: {url}")
        if url == "https://example.org/paper.pdf":
            return Response(b"%PDF-1.4 fake pdf bytes")
        raise AssertionError(f"unexpected url: {url}")

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)

    result = add_item_by_doi("10.1000/pdf")
    assert result["pdf_attached"] is True
    assert result["pdf_status"] == "attached"
    assert fake_write.attachments and fake_write.attachments[0][1] == fake_write.next_item_key


def test_add_item_by_doi_pdf_download_failure_does_not_fail_add(monkeypatch):
    monkeypatch.setenv("ZOTERO_MCP_CONTACT_EMAIL", "test@example.org")
    fake_write = FakeWriteClient()
    monkeypatch.setattr("research_mcp.zotero._write_client_from_env", lambda: fake_write)
    monkeypatch.setattr(
        "research_mcp.zotero._fetch_crossref_work",
        lambda doi: {"type": "journal-article", "title": ["T"]},
    )
    monkeypatch.setattr(
        "research_mcp.zotero._find_unpaywall_pdf_url", lambda doi: "https://example.org/broken.pdf"
    )

    def fake_urlopen(request, timeout):
        url = request.full_url
        if url.startswith("http://127.0.0.1"):
            if "/items/top" in url:
                return Response(json.dumps([]).encode())
            if url.rstrip("/").endswith(f"/items/{fake_write.next_item_key}"):
                return Response(json.dumps({"key": fake_write.next_item_key, "data": {}}).encode())
            raise AssertionError(f"unexpected local API call: {url}")
        if url == "https://example.org/broken.pdf":
            raise OSError("connection reset")
        raise AssertionError(f"unexpected url: {url}")

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)

    result = add_item_by_doi("10.1000/pdf2")
    assert result["success"] is True
    assert result["pdf_attached"] is False
    assert result["pdf_status"] == "download failed"


def test_search_refuses_out_of_range_limit():
    client = ZoteroClient()
    with pytest.raises(RefusedUnsafe, match="search limit must be between 1 and 50"):
        client.search("query", limit=0)
    with pytest.raises(RefusedUnsafe, match="search limit must be between 1 and 50"):
        client.search("query", limit=51)


def test_add_item_by_doi_does_not_create_collection_if_crossref_fails(monkeypatch):
    fake_write = FakeWriteClient()
    monkeypatch.setattr("research_mcp.zotero._write_client_from_env", lambda: fake_write)

    def fake_urlopen(request, timeout):
        url = request.full_url
        if "items/top" in url:
            return Response(json.dumps([]).encode())
        if "collections" in url:
            return Response(json.dumps([]).encode())
        if "crossref.org" in url:
            raise HTTPError(url, 404, "Not Found", None, None)
        raise AssertionError(f"unexpected url: {url}")

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)

    with pytest.raises(ZoteroItemNotFound):
        add_item_by_doi("10.1000/doesnotexist", collection="New Topic", create_collection=True)

    assert fake_write.created_collections == []


def test_add_item_by_doi_refuses_nonexistent_8char_collection_key(monkeypatch):
    def fake_urlopen(request, timeout):
        if "collections" in request.full_url:
            return Response(json.dumps([]).encode())
        raise AssertionError(f"unexpected url: {request.full_url}")

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)

    with pytest.raises(RefusedUnsafe, match="collection key 'NOTEXIST' was not found"):
        add_item_by_doi("10.1000/paper", collection="NOTEXIST")
