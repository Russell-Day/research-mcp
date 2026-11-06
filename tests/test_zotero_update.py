"""update_zotero_item: web-API metadata edits with a version precondition and a local-sync wait."""

from __future__ import annotations

import io
import json
import os

import pytest

from research_mcp.errors import RefusedUnsafe, ZoteroItemNotFound, ZoteroUnavailable
from research_mcp.zotero import update_item

KEY = "VJE7HDNH"
TYPE_FIELDS = ["title", "proceedingsTitle", "conferenceName", "volume", "pages", "url", "series", "extra", "date"]


class Response(io.BytesIO):
    status = 200

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()


class PreConditionFailedError(Exception):
    pass


class ResourceNotFoundError(Exception):
    pass


class FakeResponse:
    def __init__(self, version):
        self.headers = {"Last-Modified-Version": str(version)}


class FakeWriteClient:
    def __init__(self, data, *, conflict=False, missing=False):
        self.data = data
        self.conflict = conflict
        self.missing = missing
        self.updates: list[dict] = []
        self.request = None

    def item(self, key):
        if self.missing:
            raise ResourceNotFoundError("404")
        return {"key": key, "version": self.data["version"], "data": dict(self.data)}

    def item_type_fields(self, item_type):
        assert item_type == self.data["itemType"]
        return [{"field": name, "localized": name} for name in TYPE_FIELDS]

    def item_creator_types(self, item_type):
        return [{"creatorType": "author"}, {"creatorType": "editor"}]

    def update_item(self, payload):
        if self.conflict:
            raise PreConditionFailedError("412")
        self.updates.append(payload)
        new_version = self.data["version"] + 1
        self.data.update({k: v for k, v in payload.items() if k != "version"})
        self.data["version"] = new_version
        self.request = FakeResponse(new_version)
        return True


def base_item(**extra):
    data = {
        "key": KEY, "version": 41, "itemType": "conferencePaper",
        "title": "Fine-tuning a model", "proceedingsTitle": "Medical Imaging 2026",
        "conferenceName": "", "volume": "", "pages": "1-10", "url": "https://x",
        "series": "", "extra": "", "date": "2026",
        "creators": [{"creatorType": "author", "firstName": "J", "lastName": "Doe"}],
        "tags": [],
    }
    data.update(extra)
    return data


def install(monkeypatch, client, local_data=None, *, local_lag=0):
    """Local API mirrors `client.data` after `local_lag` polls (or a fixed `local_data`)."""
    monkeypatch.setattr("research_mcp.zotero._write_client_from_env", lambda: client)
    polls = {"n": 0}
    stale = dict(client.data)

    def fake_urlopen(request, timeout):
        url = request.full_url
        assert url.rstrip("/").endswith(f"/items/{KEY}"), url
        polls["n"] += 1
        if local_data is not None:
            data = local_data
        elif polls["n"] <= local_lag:
            data = stale
        else:
            data = client.data
        return Response(json.dumps({"key": KEY, "data": data}).encode())

    monkeypatch.setattr("research_mcp.zotero.urlopen", fake_urlopen)
    monkeypatch.setattr("research_mcp.zotero.time.sleep", lambda s: None)
    return polls


def test_update_writes_diff_with_version_precondition_and_waits_for_sync(monkeypatch):
    client = FakeWriteClient(base_item())
    polls = install(monkeypatch, client, local_lag=2)

    result = update_item(KEY, {"volume": "12345", "pages": "", "conferenceName": "SPIE Medical Imaging"})

    assert client.updates == [{
        "key": KEY, "version": 41, "volume": "12345", "pages": "", "conferenceName": "SPIE Medical Imaging",
    }]
    assert result["changed"] is True
    assert {c["field"]: (c["old"], c["new"]) for c in result["changes"]} == {
        "volume": ("", "12345"), "pages": ("1-10", ""), "conferenceName": ("", "SPIE Medical Imaging"),
    }
    assert result["previous"] == {"volume": "", "pages": "1-10", "conferenceName": ""}
    assert result["version_before"] == 41 and result["version_after"] == 42
    assert result["synced_locally"] is True
    assert polls["n"] == 3


def test_previous_values_undo_the_change(monkeypatch):
    client = FakeWriteClient(base_item())
    install(monkeypatch, client)
    first = update_item(KEY, {"volume": "12345", "pages": ""})
    second = update_item(KEY, first["previous"])
    assert client.data["volume"] == "" and client.data["pages"] == "1-10"
    assert second["previous"] == {"volume": "12345", "pages": ""}


def test_undo_restores_original_unusual_characters_exactly(monkeypatch):
    original_title = "Fine\u2011tuning a model"
    original_creators = [{"creatorType": "author", "firstName": "Jos\u00e9\u00a0M", "lastName": "Doe"}]
    client = FakeWriteClient(base_item(title=original_title, creators=original_creators))
    install(monkeypatch, client)
    first = update_item(KEY, {"title": "Fine-tuning a model", "creators": [
        {"creatorType": "author", "firstName": "Jose", "lastName": "Doe"}]})
    update_item(KEY, first["previous"], clean_text=False)
    assert client.data["title"] == original_title
    assert client.data["creators"] == original_creators


def test_surname_only_creator_matches_zotero_serialization(monkeypatch):
    client = FakeWriteClient(base_item())
    install(monkeypatch, client)
    original_update = client.update_item

    def zotero_normalizes(payload):
        original_update(payload)
        client.data["creators"] = [
            {"creatorType": c["creatorType"], "firstName": c.get("firstName", ""), "lastName": c["lastName"]}
            for c in client.data["creators"]
        ]
        return True

    client.update_item = zotero_normalizes
    result = update_item(KEY, {"creators": [{"creatorType": "author", "lastName": "Doe"}]})
    assert result["synced_locally"] is True
    # Same creators again (Zotero's shape vs the caller's) is no change.
    again = update_item(KEY, {"creators": [{"creatorType": "author", "lastName": "Doe"}]})
    assert again["changed"] is False


def test_missing_response_version_is_an_error_not_a_pass(monkeypatch):
    client = FakeWriteClient(base_item())
    install(monkeypatch, client)
    original_update = client.update_item

    def no_header(payload):
        original_update(payload)
        client.request = type("R", (), {"headers": {}})()
        return True

    client.update_item = no_header
    with pytest.raises(ZoteroUnavailable, match="no Last-Modified-Version.*clean_text=false"):
        update_item(KEY, {"volume": "12345"})


def test_missing_local_version_never_counts_as_synced(monkeypatch):
    client = FakeWriteClient(base_item())
    local = base_item(volume="12345")
    del local["version"]
    install(monkeypatch, client, local_data=local)
    clock = iter(range(0, 1000, 10))
    monkeypatch.setattr("research_mcp.zotero.time.monotonic", lambda: next(clock))
    with pytest.raises(ZoteroUnavailable, match="not synced"):
        update_item(KEY, {"volume": "12345"}, sync_timeout_seconds=30)


def test_unchanged_fields_write_nothing(monkeypatch):
    client = FakeWriteClient(base_item())
    install(monkeypatch, client)
    result = update_item(KEY, {"title": "Fine-tuning a model"})
    assert client.updates == []
    assert result["changed"] is False and result["changes"] == []
    assert result["synced_locally"] is True  # the local Zotero was checked too


def test_retry_after_sync_timeout_still_waits_for_local_sync(monkeypatch):
    # Reviewers' case: the web update succeeds (version 42) but the local Zotero stays at 41 and the
    # call times out. Repeating it finds nothing to write on the web; it must not report success
    # while the local Zotero (which renders Word citations) is still stale.
    client = FakeWriteClient(base_item())
    polls = install(monkeypatch, client, local_data=base_item())  # local stays at version 41
    clock = iter(range(0, 10000, 10))
    monkeypatch.setattr("research_mcp.zotero.time.monotonic", lambda: next(clock))
    with pytest.raises(ZoteroUnavailable, match="not synced"):
        update_item(KEY, {"volume": "12345"}, sync_timeout_seconds=30)
    assert client.data["version"] == 42
    reads = polls["n"]
    with pytest.raises(ZoteroUnavailable, match=r"already has these values \(version 42\).*not synced"):
        update_item(KEY, {"volume": "12345"}, sync_timeout_seconds=30)
    assert polls["n"] > reads and len(client.updates) == 1
    # Once the local Zotero has synced, the retry succeeds without writing again.
    install(monkeypatch, client)
    result = update_item(KEY, {"volume": "12345"}, sync_timeout_seconds=30)
    assert result["changed"] is False and result["synced_locally"] is True and len(client.updates) == 1


def test_dry_run_returns_diff_without_writing(monkeypatch):
    client = FakeWriteClient(base_item())
    install(monkeypatch, client)
    result = update_item(KEY, {"volume": "12345"}, dry_run=True)
    assert client.updates == []
    assert result["dry_run"] is True and result["changes"][0]["new"] == "12345"
    assert result["synced_locally"] is None


def test_text_is_cleaned_u2011_becomes_hyphen(monkeypatch):
    client = FakeWriteClient(base_item())
    install(monkeypatch, client)
    result = update_item(KEY, {"title": "Fine‑tuning a model v2"})
    assert client.updates[0]["title"] == "Fine-tuning a model v2"
    assert result["cleaned"] and result["cleaned"][0]["field"] == "title"


@pytest.mark.parametrize("field", ["itemType", "key", "version", "collections", "dateModified", "deleted"])
def test_identity_fields_are_refused(monkeypatch, field):
    client = FakeWriteClient(base_item())
    install(monkeypatch, client)
    with pytest.raises(RefusedUnsafe, match="never changes"):
        update_item(KEY, {field: "x"})
    assert client.updates == []


def test_unknown_field_names_the_valid_fields(monkeypatch):
    client = FakeWriteClient(base_item())
    install(monkeypatch, client)
    with pytest.raises(RefusedUnsafe, match=r"not a field of a conferencePaper item: \['publicationTitle'\]"):
        update_item(KEY, {"publicationTitle": "Proc SPIE"})
    assert client.updates == []


def test_non_string_field_value_refused(monkeypatch):
    client = FakeWriteClient(base_item())
    install(monkeypatch, client)
    with pytest.raises(RefusedUnsafe, match="takes a string"):
        update_item(KEY, {"volume": 12345})


def test_creators_replace_whole_list_and_types_are_checked(monkeypatch):
    client = FakeWriteClient(base_item())
    install(monkeypatch, client)
    new = [{"creatorType": "author", "firstName": "Jane", "lastName": "Doe"},
           {"creatorType": "editor", "name": "SPIE"}]
    update_item(KEY, {"creators": new})
    assert client.updates[0]["creators"] == new
    with pytest.raises(RefusedUnsafe, match="creatorType 'translator'"):
        update_item(KEY, {"creators": [{"creatorType": "translator", "name": "X"}]})
    with pytest.raises(RefusedUnsafe, match="either lastName/firstName or a single name"):
        update_item(KEY, {"creators": [{"creatorType": "author", "name": "X", "lastName": "Y"}]})


def test_trashed_item_refused(monkeypatch):
    client = FakeWriteClient(base_item(deleted=1))
    install(monkeypatch, client)
    with pytest.raises(RefusedUnsafe, match="in the trash"):
        update_item(KEY, {"volume": "1"})
    assert client.updates == []


def test_note_item_refused(monkeypatch):
    client = FakeWriteClient(base_item(itemType="note"))
    install(monkeypatch, client)
    with pytest.raises(RefusedUnsafe, match="is a note"):
        update_item(KEY, {"note": "x"})


def test_missing_item_raises_not_found(monkeypatch):
    client = FakeWriteClient(base_item(), missing=True)
    install(monkeypatch, client)
    with pytest.raises(ZoteroItemNotFound):
        update_item(KEY, {"volume": "1"})


def test_conflict_with_newer_server_version_refuses(monkeypatch):
    client = FakeWriteClient(base_item(), conflict=True)
    install(monkeypatch, client)
    with pytest.raises(RefusedUnsafe, match="changed on the server since version 41"):
        update_item(KEY, {"volume": "1"})


def test_sync_timeout_is_an_error_naming_the_undo(monkeypatch):
    client = FakeWriteClient(base_item())
    install(monkeypatch, client, local_data=base_item())  # local never catches up
    clock = iter(range(0, 1000, 10))
    monkeypatch.setattr("research_mcp.zotero.time.monotonic", lambda: next(clock))
    with pytest.raises(ZoteroUnavailable, match=r"succeeded \(version 42\).*not synced.*clean_text=false.*\"volume\": \"\""):
        update_item(KEY, {"volume": "12345"}, sync_timeout_seconds=30)


def test_bad_key_refused_before_any_call(monkeypatch):
    monkeypatch.setattr(
        "research_mcp.zotero._write_client_from_env",
        lambda: (_ for _ in ()).throw(AssertionError("no client")),
    )
    with pytest.raises(RefusedUnsafe, match="not a Zotero item key"):
        update_item("../x", {"volume": "1"})


# -- live: a throwaway item in "research-mcp tests", trashed afterwards ----------

LIVE = os.environ.get("RESEARCH_MCP_ZOTERO_LIVE") == "1"


@pytest.mark.skipif(not LIVE, reason="set RESEARCH_MCP_ZOTERO_LIVE=1 (web API + local Zotero running)")
def test_live_update_on_throwaway_item():
    from research_mcp.zotero import ZoteroClient, _write_client_from_env

    client = _write_client_from_env()
    name = "research-mcp tests"
    collections = [c for c in client.everything(client.collections()) if c["data"]["name"] == name]
    if collections:
        collection_key = collections[0]["key"]
    else:
        created = client.create_collections([{"name": name}])
        collection_key = next(iter(created["success"].values()))
    template = client.item_template("conferencePaper")
    template.update({
        "title": "research-mcp test item, safe to delete",
        "proceedingsTitle": "Test proceedings", "pages": "1-2",
        "collections": [collection_key],
    })
    created = client.create_items([template])
    key = next(iter(created["success"].values()))
    try:
        local = ZoteroClient()
        result = update_item(
            key,
            {"title": "research‑mcp test item, safe to delete (edited)", "volume": "12345", "pages": ""},
            sync_timeout_seconds=120,
        )
        assert result["synced_locally"] is True, result
        data = local.item(key)
        assert data["title"] == "research-mcp test item, safe to delete (edited)"
        assert data["volume"] == "12345" and data.get("pages", "") == ""

        # A stale version conflicts and writes nothing.
        stale = dict(client.item(key)["data"])
        client.update_item({"key": key, "version": stale["version"], "extra": "bump"})
        with pytest.raises(Exception) as excinfo:
            client.update_item({"key": key, "version": stale["version"], "extra": "stale"})
        assert type(excinfo.value).__name__ == "PreConditionFailedError"
        assert client.item(key)["data"]["extra"] == "bump"

        undo = update_item(key, result["previous"], clean_text=False, sync_timeout_seconds=120)
        assert undo["synced_locally"] is True
        data = local.item(key)
        assert data.get("volume", "") == "" and data["pages"] == "1-2"
    finally:
        # Move to the trash (not a permanent delete); pyzotero has no trash call.
        fresh = client.item(key)
        response = client._write(
            "PATCH",
            url=f"{client.endpoint}/{client.library_type}/{client.library_id}/items/{key}",
            headers={"If-Unmodified-Since-Version": str(fresh["version"])},
            json={"deleted": 1},
        )
        assert response.status_code == 204, response.text
