"""Small Zotero local-API client; raw CSL never reaches the MCP surface."""

from __future__ import annotations

from dataclasses import dataclass
import difflib
import json
import os
import re
import tempfile
import time
from html.parser import HTMLParser
from pathlib import Path
from typing import Any
from urllib.error import HTTPError
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen

from . import config
from .errors import RefusedUnsafe, ZoteroItemNotFound, ZoteroUnavailable

_COLLECTION_KEY_RE = re.compile(r"^[A-Z0-9]{8}$")

CROSSREF_TYPE_MAP = {
    "journal-article": "journalArticle",
    "book": "book",
    "book-chapter": "bookSection",
    "proceedings-article": "conferencePaper",
    "report": "report",
    "dissertation": "thesis",
    "posted-content": "preprint",
    "monograph": "book",
    "reference-entry": "encyclopediaArticle",
    "dataset": "document",
    "peer-review": "document",
    "edited-book": "book",
    "standard": "document",
}


def normalize_doi(doi: str) -> str:
    """Lowercase, strip common prefixes, and reject anything that is not a DOI."""

    text = str(doi).strip().lower()
    for prefix in ("https://doi.org/", "http://doi.org/", "https://dx.doi.org/",
                   "http://dx.doi.org/", "doi:", "doi "):
        if text.startswith(prefix):
            text = text[len(prefix):].strip()
    if not text.startswith("10.") or "/" not in text:
        raise RefusedUnsafe(f"not a well-formed DOI: {doi!r}")
    return text


@dataclass(frozen=True)
class Candidate:
    key: str
    title: str
    authors: str
    year: str


class ZoteroClient:
    def __init__(self, base_url: str = "http://127.0.0.1:23119", *, timeout: float = 15.0):
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.headers = {
            "Zotero-Allowed-Request": "1",
            "User-Agent": "research-mcp/0.1",
            "Accept": "application/json",
        }

    def _get(self, path: str) -> Any:
        request = Request(self.base_url + path, headers=self.headers)
        try:
            with urlopen(request, timeout=self.timeout) as response:
                return json.loads(response.read().decode("utf-8"))
        except Exception as exc:
            raise ZoteroUnavailable(f"Zotero local API request failed for {path}: {exc}") from exc

    def ping(self) -> bool:
        request = Request(self.base_url + "/connector/ping", headers=self.headers)
        try:
            with urlopen(request, timeout=self.timeout) as response:
                return response.status == 200
        except Exception as exc:
            raise ZoteroUnavailable(f"Zotero connector is unavailable: {exc}") from exc

    def item(self, key: str) -> dict[str, Any]:
        """Raw library item for one key, including its `deleted` (trash) flag."""

        payload = self._get(f"/api/users/0/items/{quote(key)}")
        if not isinstance(payload, dict):
            raise ZoteroUnavailable(f"unexpected item response for {key!r}")
        data = payload.get("data", payload)
        if not isinstance(data, dict):
            raise ZoteroUnavailable(f"unexpected item payload for {key!r}")
        return data

    def csl(self, key: str) -> dict[str, Any]:
        # Search hides trashed items but a direct key fetch does not, so without
        # this guard a stale key silently produces a live Word citation pointing
        # at an item that vanishes when the user empties the trash.
        if self.item(key).get("deleted"):
            raise RefusedUnsafe(
                f"Zotero item {key!r} is in the trash; restore it before citing it"
            )
        data = self._get(f"/api/users/0/items/{quote(key)}?format=csljson")
        if not isinstance(data, list) or len(data) != 1 or not isinstance(data[0], dict):
            raise ZoteroUnavailable(f"unexpected CSL response for item {key!r}")
        return data[0]

    def find_by_doi(self, doi: str) -> str | None:
        """Key of the one non-trashed library item with this DOI, if any.

        The `q`/`qmode=everything` search endpoint does not index the DOI field
        -- searching a live item's exact DOI returns nothing -- so DOI lookup
        must scan top-level items instead. Using search here would miss every
        existing item and create duplicates.
        """

        wanted = normalize_doi(doi)
        matches: list[str] = []
        for entry in self.top_level():
            data = entry.get("data", entry) if isinstance(entry, dict) else {}
            if not isinstance(data, dict) or data.get("deleted"):
                continue
            raw = data.get("DOI") or ""
            if not raw:
                continue
            try:
                if normalize_doi(str(raw)) == wanted:
                    matches.append(str(data.get("key") or entry.get("key") or ""))
            except RefusedUnsafe:
                continue
        matches = [key for key in matches if key]
        if len(matches) > 1:
            raise RefusedUnsafe(
                f"DOI {wanted!r} matches multiple Zotero items {sorted(matches)}; "
                "resolve the duplicates before citing"
            )
        return matches[0] if matches else None

    def top_level(self) -> list[dict[str, Any]]:
        """Every top-level library item, for reference resolution."""
        collected: list[dict[str, Any]] = []
        start = 0
        while True:
            page = self._get(f"/api/users/0/items/top?limit=100&start={start}")
            if not isinstance(page, list) or not page:
                break
            collected.extend(page)
            if len(page) < 100:
                break
            start += len(page)
        return collected

    def collections(self) -> list[dict[str, Any]]:
        """Every collection (folder) in the personal library, with key/name/parent."""
        collected: list[dict[str, Any]] = []
        start = 0
        while True:
            page = self._get(f"/api/users/0/collections?limit=100&start={start}")
            if not isinstance(page, list) or not page:
                break
            for entry in page:
                data = entry.get("data", entry) if isinstance(entry, dict) else {}
                if not isinstance(data, dict):
                    continue
                collected.append({
                    "key": str(data.get("key") or entry.get("key") or ""),
                    "name": str(data.get("name") or ""),
                    "parent": data.get("parentCollection") or None,
                })
            if len(page) < 100:
                break
            start += len(page)
        return collected

    def find_collection_key(self, name_or_key: str) -> str | None:
        """Collection key for a name or key, or None if no collection has that name.

        Raises RefusedUnsafe if the name matches more than one collection --
        nested collections can share a name, so a silent first-match would
        risk filing into the wrong folder.

        A value matching the Zotero key shape is returned unchanged and is not
        checked against the live collection list, so a syntactically valid but
        nonexistent key passes through here. Callers that write must confirm
        the key exists themselves; `set_project_collection` does.
        """
        candidate = name_or_key.strip()
        if _COLLECTION_KEY_RE.match(candidate):
            return candidate
        matches = [c for c in self.collections() if c["name"] == candidate]
        if not matches:
            return None
        if len(matches) > 1:
            raise RefusedUnsafe(
                f"collection name {candidate!r} is ambiguous ({len(matches)} matches); "
                "pass the collection key instead"
            )
        return matches[0]["key"]

    def resolve_collection(self, name_or_key: str) -> str:
        """Resolve a collection name to its key, or raise with the available names.

        A value already matching the key shape passes through unverified; see
        `find_collection_key`.
        """
        key = self.find_collection_key(name_or_key)
        if key is None:
            collections = self.collections()
            available = sorted({c["name"] for c in collections})
            suggestions = _collection_suggestions(name_or_key, collections)
            hint = f" did you mean {', '.join(suggestions)}?" if suggestions else ""
            raise RefusedUnsafe(
                f"no Zotero collection named {name_or_key!r};{hint} available: {available}"
            )
        return key

    def search(
        self, query: str, *, limit: int = 10, collection_key: str | None = None
    ) -> tuple[Candidate, ...]:
        if not (1 <= limit <= 50):
            raise RefusedUnsafe(f"search limit must be between 1 and 50, got {limit}")
        params = urlencode({"q": query, "qmode": "everything", "limit": limit})
        prefix = f"/api/users/0/collections/{quote(collection_key)}" if collection_key else "/api/users/0"
        data = self._get(f"{prefix}/items?{params}")
        candidates: list[Candidate] = []
        seen: set[str] = set()
        for item in data if isinstance(data, list) else []:
            key = str(item.get("key") or item.get("data", {}).get("key") or "")
            metadata = item.get("data", item)
            if not key or key in seen or metadata.get("itemType") == "attachment":
                continue
            seen.add(key)
            creators = metadata.get("creators", [])
            authors = ", ".join(
                creator.get("lastName") or creator.get("name") or ""
                for creator in creators[:3]
            ).strip(", ")
            date = str(metadata.get("date", ""))
            candidates.append(
                Candidate(
                    key=key,
                    title=str(metadata.get("title", "")),
                    authors=authors,
                    year=date[:4] if len(date) >= 4 else date,
                )
            )
        return tuple(candidates)


def _collection_suggestions(name: str, collections: list[dict[str, Any]]) -> list[str]:
    """Up to 3 close-name matches for `name`, each labeled with its parent.

    Labeled per occurrence, not deduped by name -- two collections that share
    a name under different parents (nested collections can) are both worth
    surfacing so the caller can tell them apart.
    """
    by_key = {c["key"]: c["name"] for c in collections}
    close_names = difflib.get_close_matches(
        name, [c["name"] for c in collections], n=3, cutoff=0.6
    )
    labeled: list[str] = []
    for close_name in close_names:
        for entry in collections:
            if entry["name"] != close_name:
                continue
            parent_name = by_key.get(entry["parent"]) if entry["parent"] else None
            suffix = f" (under {parent_name!r})" if parent_name else " (top-level)"
            labeled.append(f"{entry['name']!r}{suffix}")
    return labeled


def find_project_collection_details(
    doc_path: str, *, max_depth: int = 8
) -> tuple[str, Path] | None:
    """Return the inherited marker value and exact marker path, if present."""
    directory = Path(doc_path).resolve().parent
    for _ in range(max_depth):
        marker = directory / ".zotero-collection"
        if marker.is_file():
            value = _marker_value(marker)
            if value:
                return (value, marker.resolve())
            # A blank marker is not a signal to stop -- keep walking so a
            # parent's real marker can still be inherited.
        parent = directory.parent
        if parent == directory:
            break
        directory = parent
    return None


def find_project_collection(doc_path: str, *, max_depth: int = 8) -> str | None:
    """Walk up from `doc_path`'s folder for a `.zotero-collection` marker.

    The marker's first line is the target collection's name (or key). Lets a
    project auto-scope to its own Zotero folder without a param on every call.
    """
    details = find_project_collection_details(doc_path, max_depth=max_depth)
    return details[0] if details else None


def _marker_folder(doc_or_dir_path: str | Path) -> Path:
    """Resolve a marker target as an existing directory or a file's parent."""
    candidate = Path(doc_or_dir_path).expanduser().resolve()
    folder = candidate if candidate.is_dir() else candidate.parent
    if not folder.is_dir():
        raise RefusedUnsafe(f"marker target folder does not exist: {folder}")
    return folder


def _marker_value(marker: Path) -> str | None:
    lines = marker.read_text(encoding="utf-8").splitlines()
    return lines[0].strip() if lines and lines[0].strip() else None


def write_project_collection(
    doc_or_dir_path: str | Path,
    name: str,
    *,
    replace_existing: bool = False,
) -> dict[str, Any]:
    """Write one exact collection name to a project's marker file.

    The target is the supplied directory itself when it exists, otherwise the
    parent directory of the supplied document path. Existing markers are not
    overwritten unless explicitly requested."""
    value = str(name).strip()
    if not value:
        raise RefusedUnsafe("collection name must not be empty")
    if "\r" in value or "\n" in value:
        raise RefusedUnsafe("collection name must fit on one marker-file line")

    folder = _marker_folder(doc_or_dir_path)
    marker = folder / ".zotero-collection"
    previous_value: str | None = None
    existed = marker.exists()
    if existed:
        if not marker.is_file():
            raise RefusedUnsafe(f"marker path is not a file: {marker}")
        previous_value = _marker_value(marker)
        if not replace_existing:
            description = (
                f"current value is {previous_value!r}"
                if previous_value is not None
                else "it exists but is blank"
            )
            raise RefusedUnsafe(
                f"project collection marker already exists at {marker}; "
                f"{description}; pass replace_existing=true to replace it"
            )

    temporary: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            dir=folder,
            prefix=".zotero-collection.",
            suffix=".tmp",
            delete=False,
        ) as handle:
            temporary = Path(handle.name)
            handle.write(value + "\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, marker)
    finally:
        if temporary is not None and temporary.exists():
            try:
                temporary.unlink()
            except OSError:
                pass

    actual = _marker_value(marker)
    if actual != value:
        raise ZoteroUnavailable(
            f"project collection marker did not read back as {value!r}: {marker}"
        )
    return {
        "marker_path": str(marker),
        "previous_value": previous_value,
        "replaced": existed,
    }


class _TextOnlyParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.chunks: list[str] = []

    def handle_data(self, data: str) -> None:
        self.chunks.append(data)


def _strip_markup(text: str) -> str:
    """CrossRef abstracts arrive as JATS/HTML; keep only the text."""
    parser = _TextOnlyParser()
    parser.feed(text)
    return " ".join("".join(parser.chunks).split())


def _write_client_from_env():
    """A pyzotero client authorized for writes, per `ZOTERO_API_KEY`/`ZOTERO_LIBRARY_ID`.

    Both come from the environment or the config file (`config.config_path()`).

    The local connector API (everything else in this module) cannot write --
    see reference/zotero-mcp's `_get_write_client`, which raises for the same
    reason. Writes go through the web API instead.
    """
    from pyzotero import zotero as pyzotero_zotero

    library_id = config.get("ZOTERO_LIBRARY_ID", "")
    api_key = config.get("ZOTERO_API_KEY", "")
    library_type = config.get("ZOTERO_LIBRARY_TYPE", "user")
    missing = [
        name for name, value in (("ZOTERO_LIBRARY_ID", library_id), ("ZOTERO_API_KEY", api_key))
        if not value
    ]
    if missing:
        raise ZoteroUnavailable(
            f"Zotero write access needs {', '.join(missing)} set in the environment or in {config.config_path()}"
        )
    return pyzotero_zotero.Zotero(library_id=library_id, library_type=library_type, api_key=api_key)


def _fetch_crossref_work(doi: str) -> dict[str, Any]:
    url = f"https://api.crossref.org/works/{quote(doi)}"
    contact_email = config.get("ZOTERO_MCP_CONTACT_EMAIL", "")
    if contact_email:
        url += f"?mailto={quote(contact_email)}"
    request = Request(url, headers={"User-Agent": "research-mcp/0.1", "Accept": "application/json"})
    try:
        with urlopen(request, timeout=15.0) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        if exc.code == 404:
            raise ZoteroItemNotFound(f"DOI not found on CrossRef: {doi}") from exc
        raise ZoteroUnavailable(f"CrossRef request failed for {doi}: {exc}") from exc
    except Exception as exc:
        raise ZoteroUnavailable(f"CrossRef request failed for {doi}: {exc}") from exc
    message = payload.get("message")
    if not isinstance(message, dict):
        raise ZoteroUnavailable(f"unexpected CrossRef response for {doi}")
    return message


def _find_unpaywall_pdf_url(doi: str) -> str | None:
    """Best open-access PDF URL for `doi`, or None if unfound/unconfigured.

    Requires ZOTERO_MCP_CONTACT_EMAIL -- Unpaywall requires ?email= on every
    request, and that must not silently default to the operator's personal
    address.
    """
    contact_email = config.get("ZOTERO_MCP_CONTACT_EMAIL", "")
    if not contact_email:
        return None
    url = f"https://api.unpaywall.org/v2/{quote(doi)}?email={quote(contact_email)}"
    request = Request(url, headers={"User-Agent": "research-mcp/0.1", "Accept": "application/json"})
    try:
        with urlopen(request, timeout=15.0) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except Exception:
        return None
    location = payload.get("best_oa_location")
    if not isinstance(location, dict):
        return None
    pdf_url = location.get("url_for_pdf") or location.get("url")
    return str(pdf_url) if pdf_url else None


def _first_author(creators: list[dict[str, Any]]) -> str:
    for creator in creators:
        if creator.get("creatorType") == "author":
            return str(creator.get("lastName") or creator.get("name") or "")
    return ""


def _create_collection(write_client, name: str) -> str:
    created = write_client.create_collections([{"name": name}])
    success = created.get("success") if isinstance(created, dict) else None
    if not success:
        raise ZoteroUnavailable(f"failed to create Zotero collection {name!r}: {created}")
    return next(iter(success.values()))


def add_item_by_doi(
    doi: str,
    *,
    path: str | None = None,
    collection: str | None = None,
    create_collection: bool = False,
) -> dict[str, Any]:
    """DOI in, full CrossRef metadata + best-effort OA PDF Zotero item out.

    Returns the existing item (created=False) instead of a duplicate when the
    DOI is already in the library. Files into `collection`, or the collection
    named by a `.zotero-collection` marker found walking up from `path`, if
    either resolves; otherwise the item is left unfiled. The marker is a plain
    text file whose first line is the exact collection name, with no quotes or
    trailing punctuation.
    """
    normalized = normalize_doi(doi)
    read_client = ZoteroClient()

    marker_details = find_project_collection_details(path) if path else None
    target_name = collection or (marker_details[0] if marker_details else None)
    marker_path = (
        str(marker_details[1]) if marker_details and collection is None else None
    )
    target_key: str | None = None
    if target_name is not None:
        target_key = read_client.find_collection_key(target_name)
        if target_key is not None and _COLLECTION_KEY_RE.match(target_name.strip()):
            collections = read_client.collections()
            if not any(c["key"] == target_key for c in collections):
                raise RefusedUnsafe(
                    f"collection key {target_key!r} was not found in the live collection list"
                )
        if target_key is None:
            collections = read_client.collections()
            suggestions = _collection_suggestions(target_name, collections)
            if not create_collection:
                available = sorted({c["name"] for c in collections})
                hint = f" did you mean {', '.join(suggestions)}?" if suggestions else ""
                raise RefusedUnsafe(
                    f"no Zotero collection named {target_name!r};{hint} available: "
                    f"{available}. Retry with create_collection=true to create it."
                )
            if suggestions:
                # A typo here would silently create a near-duplicate folder.
                # Refuse and require an exact retry -- either the existing
                # close match (files into it, nothing created) or, if a
                # genuinely new collection is intended, confirmation that the
                # name isn't a typo is the only check an MCP server can
                # actually enforce.
                raise RefusedUnsafe(
                    f"no Zotero collection named {target_name!r}, and create_collection=true "
                    f"would create one, but it is close to {', '.join(suggestions)}; this "
                    "tool never auto-creates a near-duplicate. If this was a typo, retry "
                    "with that exact existing name. If a new, similarly-named collection "
                    "is genuinely intended, create it once in Zotero directly, then retry "
                    "this call -- it will resolve to the existing key instead of creating"
                )

    existing_key = read_client.find_by_doi(normalized)
    if existing_key is not None:
        data = read_client.item(existing_key)
        result = {
            "success": True,
            "created": False,
            "item_key": existing_key,
            "doi": normalized,
            "title": str(data.get("title", "")),
            "first_author": _first_author(data.get("creators", [])),
            "abstract": str(data.get("abstractNote", "")),
            "pdf_attached": False,
            "pdf_status": "skipped (existing item)",
            "visible_locally": True,
            "collection": None,
            "collection_marker_path": marker_path,
        }
        if target_name is not None:
            write_client = _write_client_from_env()
            if target_key is None:
                target_key = _create_collection(write_client, target_name)
            write_client.addto_collection(target_key, write_client.item(existing_key))
            result["collection"] = target_name
        return result

    # Fetch CrossRef metadata first before creating collections or write client
    work = _fetch_crossref_work(normalized)

    write_client = _write_client_from_env()

    if target_name is not None and target_key is None:
        target_key = _create_collection(write_client, target_name)

    zot_type = CROSSREF_TYPE_MAP.get(str(work.get("type", "")), "document")
    item_data = dict(write_client.item_template(zot_type))

    titles = work.get("title") or []
    if titles and "title" in item_data:
        item_data["title"] = titles[0]

    creators: list[dict[str, str]] = []
    for author in work.get("author", []) or []:
        if "family" in author:
            creators.append({
                "creatorType": "author", "firstName": author.get("given", ""), "lastName": author["family"],
            })
        elif "name" in author:
            creators.append({"creatorType": "author", "name": author["name"]})
    for editor in work.get("editor", []) or []:
        if "family" in editor:
            creators.append({
                "creatorType": "editor", "firstName": editor.get("given", ""), "lastName": editor["family"],
            })
        elif "name" in editor:
            creators.append({"creatorType": "editor", "name": editor["name"]})
    if creators and "creators" in item_data:
        item_data["creators"] = creators

    date_parts = ((work.get("published") or work.get("created") or {}).get("date-parts") or [[]])[0]
    if date_parts and "date" in item_data:
        item_data["date"] = "-".join(str(part) for part in date_parts)

    abstract = _strip_markup(str(work.get("abstract", "")))
    field_values = {
        "DOI": normalized,
        "url": str(work.get("URL", "")),
        "volume": str(work.get("volume", "")),
        "issue": str(work.get("issue", "")),
        "pages": str(work.get("page", "")),
        "publisher": str(work.get("publisher", "")),
        "ISSN": (work.get("ISSN") or [""])[0],
        "publicationTitle": (work.get("container-title") or [""])[0],
        "abstractNote": abstract,
    }
    for field, value in field_values.items():
        if field in item_data and value:
            item_data[field] = value

    if target_key:
        item_data["collections"] = [target_key]

    created_result = write_client.create_items([item_data])
    success = created_result.get("success") if isinstance(created_result, dict) else None
    if not success:
        raise ZoteroUnavailable(f"failed to create Zotero item for {normalized}: {created_result}")
    item_key = next(iter(success.values()))

    pdf_attached = False
    pdf_status = "no contact email configured"
    if config.get("ZOTERO_MCP_CONTACT_EMAIL", ""):
        pdf_url = _find_unpaywall_pdf_url(normalized)
        if pdf_url is None:
            pdf_status = "not found"
        else:
            tmp_path: str | None = None
            try:
                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                    tmp_path = tmp.name
                    request = Request(pdf_url, headers={"User-Agent": "research-mcp/0.1"})
                    with urlopen(request, timeout=30.0) as response:
                        tmp.write(response.read())
                write_client.attachment_simple([tmp_path], item_key)
                pdf_attached = True
                pdf_status = "attached"
            except Exception:
                pdf_status = "download failed"
            finally:
                if tmp_path:
                    try:
                        os.unlink(tmp_path)
                    except OSError:
                        pass

    visible_locally = False
    for _ in range(5):
        try:
            read_client.item(item_key)
            visible_locally = True
            break
        except ZoteroUnavailable:
            time.sleep(1.0)

    return {
        "success": True,
        "created": True,
        "item_key": item_key,
        "doi": normalized,
        "title": str(item_data.get("title", "")),
        "first_author": _first_author(creators),
        "abstract": abstract,
        "pdf_attached": pdf_attached,
        "pdf_status": pdf_status,
        "visible_locally": visible_locally,
        "collection": target_name if target_key else None,
        "collection_marker_path": marker_path,
    }


# -- update an existing item -----------------------------------------------------

# Fields a caller may never set through update_item: identity and bookkeeping.
_UPDATE_REFUSED_FIELDS = frozenset({
    "key", "version", "itemType", "dateAdded", "dateModified", "deleted",
    "relations", "parentItem", "collections",
})
# Fields accepted besides the item type's own fields (full lists replace).
_UPDATE_LIST_FIELDS = frozenset({"creators", "tags"})
_UPDATE_REFUSED_TYPES = frozenset({"note", "attachment", "annotation"})


def _clean_value(value: Any, field: str, cleaned: list[dict[str, Any]]) -> Any:
    """Run text through textclean, recording what changed under `field`."""
    from . import textclean

    if isinstance(value, str):
        out, changes = textclean.clean(value)
        if changes:
            cleaned.append({"field": field, "changes": changes})
        return out
    if isinstance(value, list):
        return [_clean_value(entry, f"{field}[{i}]", cleaned) for i, entry in enumerate(value)]
    if isinstance(value, dict):
        return {k: _clean_value(v, f"{field}.{k}", cleaned) for k, v in value.items()}
    return value


def _validate_update(write_client, item_type: str, fields: dict[str, Any]) -> None:
    if not fields:
        raise RefusedUnsafe("fields must name at least one Zotero field to change")
    refused = sorted(set(fields) & _UPDATE_REFUSED_FIELDS)
    if refused:
        raise RefusedUnsafe(
            f"update_zotero_item never changes {refused}; only an item's metadata "
            "fields, creators and tags"
        )
    valid = {str(entry.get("field")) for entry in write_client.item_type_fields(item_type)}
    unknown = sorted(set(fields) - valid - _UPDATE_LIST_FIELDS)
    if unknown:
        raise RefusedUnsafe(
            f"not a field of a {item_type} item: {unknown}; valid fields: {sorted(valid)}"
        )
    for name, value in fields.items():
        if name in _UPDATE_LIST_FIELDS:
            if not isinstance(value, list) or not all(isinstance(e, dict) for e in value):
                raise RefusedUnsafe(f"{name} must be the full list of {name} objects")
        elif not isinstance(value, str):
            raise RefusedUnsafe(f"field {name!r} takes a string ('' clears it), got {type(value).__name__}")
    creators = fields.get("creators")
    if creators:
        allowed = {str(e.get("creatorType")) for e in write_client.item_creator_types(item_type)}
        for index, creator in enumerate(creators):
            kind = creator.get("creatorType")
            if kind not in allowed:
                raise RefusedUnsafe(
                    f"creators[{index}].creatorType {kind!r} is not valid for a {item_type} "
                    f"item; valid: {sorted(allowed)}"
                )
            has_split = "lastName" in creator or "firstName" in creator
            if has_split == ("name" in creator):
                raise RefusedUnsafe(
                    f"creators[{index}] needs either lastName/firstName or a single name"
                )
    for index, tag in enumerate(fields.get("tags") or []):
        if not isinstance(tag.get("tag"), str) or not tag["tag"].strip():
            raise RefusedUnsafe(f"tags[{index}] needs a non-empty 'tag' string")


def _normalized(name: str, value: Any) -> Any:
    """A field value in the shape Zotero stores it, for comparisons.

    Zotero serializes a two-field creator with both names (an omitted
    firstName comes back as ""), and a tag without a type as type 0.
    """
    if name == "creators":
        out = []
        for creator in value or []:
            if "name" in creator:
                out.append({"creatorType": creator.get("creatorType"), "name": creator.get("name", "")})
            else:
                out.append({
                    "creatorType": creator.get("creatorType"),
                    "firstName": creator.get("firstName", ""),
                    "lastName": creator.get("lastName", ""),
                })
        return out
    if name == "tags":
        return sorted((t.get("tag"), t.get("type", 0)) for t in value or [])
    return value if value is not None else ""


def _local_matches(data: dict[str, Any], wanted: dict[str, Any], version: int) -> bool:
    local_version = data.get("version")
    if not isinstance(local_version, int) or local_version < version:
        return False
    for name, value in wanted.items():
        current = data.get(name, [] if name in _UPDATE_LIST_FIELDS else "")
        if _normalized(name, current) != _normalized(name, value):
            return False
    return True


def update_item(
    key: str,
    fields: dict[str, Any],
    *,
    dry_run: bool = False,
    clean_text: bool = True,
    sync_timeout_seconds: float = 60.0,
    poll_seconds: float = 1.0,
) -> dict[str, Any]:
    """Change an existing Zotero item's metadata through the web API.

    References are fixed in Zotero, then refreshed in Word -- never by editing
    a document's reference-list text (a refresh overwrites hand edits). `fields` maps
    Zotero field names to new values: strings for ordinary fields ("" clears
    one), full lists for `creators` and `tags`. The write is conditional on the
    item's current web version, so a change made elsewhere since the read fails
    loudly instead of being overwritten. After the write, waits until the local
    Zotero (which renders Word citations) has synced the new values; a timeout
    is an error, never a pass. Returns the diff and the previous values, which
    undo the change when passed back as `fields` with `clean_text=False` (so
    the original text is restored exactly, unusual characters included).
    """
    key = str(key).strip()
    if not _COLLECTION_KEY_RE.match(key):
        raise RefusedUnsafe(f"not a Zotero item key: {key!r}")
    if not isinstance(fields, dict):
        raise RefusedUnsafe("fields must be an object of Zotero field names to values")

    write_client = _write_client_from_env()
    try:
        current = write_client.item(key)
    except Exception as exc:
        if type(exc).__name__ == "ResourceNotFoundError":
            raise ZoteroItemNotFound(f"no Zotero item {key!r} in the web library") from exc
        raise ZoteroUnavailable(f"Zotero web API read failed for {key!r}: {exc}") from exc
    data = current.get("data", current) if isinstance(current, dict) else {}
    if not isinstance(data, dict) or "itemType" not in data:
        raise ZoteroUnavailable(f"unexpected Zotero web API item for {key!r}")
    item_type = str(data["itemType"])
    if item_type in _UPDATE_REFUSED_TYPES:
        raise RefusedUnsafe(f"Zotero item {key!r} is a {item_type}, not a reference")
    if data.get("deleted"):
        raise RefusedUnsafe(f"Zotero item {key!r} is in the trash; restore it before editing it")
    version = data.get("version", current.get("version"))
    if not isinstance(version, int):
        raise ZoteroUnavailable(f"Zotero web API gave no version for {key!r}")

    _validate_update(write_client, item_type, fields)
    cleaned: list[dict[str, Any]] = []
    if clean_text:
        wanted = {name: _clean_value(value, name, cleaned) for name, value in fields.items()}
    else:
        wanted = dict(fields)

    changes: list[dict[str, Any]] = []
    previous: dict[str, Any] = {}
    for name, value in wanted.items():
        old = data.get(name, [] if name in _UPDATE_LIST_FIELDS else "")
        if _normalized(name, old) != _normalized(name, value):
            changes.append({"field": name, "old": old, "new": value})
            previous[name] = old
    result: dict[str, Any] = {
        "success": True,
        "item_key": key,
        "item_type": item_type,
        "title": str(data.get("title", "")),
        "changed": bool(changes),
        "dry_run": dry_run,
        "changes": changes,
        "previous": previous,
        "cleaned": cleaned,
        "version_before": version,
        "version_after": version,
        "synced_locally": None,
        "sync_seconds": None,
    }
    if dry_run:
        return result
    if not changes:
        # Nothing to write on the web, but a retry after a sync timeout lands here too: the local
        # Zotero (which renders Word citations) must still show these values before this succeeds.
        _wait_for_local_sync(
            key, wanted, version, result, sync_timeout_seconds, poll_seconds,
            f"The web item {key!r} already has these values (version {version}), but",
            "",
        )
        return result

    payload = {"key": key, "version": version, **{c["field"]: c["new"] for c in changes}}
    try:
        write_client.update_item(payload)
    except Exception as exc:
        if type(exc).__name__ in {"PreConditionFailedError", "ConflictError"}:
            raise RefusedUnsafe(
                f"Zotero item {key!r} changed on the server since version {version} was "
                "read; nothing was written. Read it again and retry."
            ) from exc
        raise ZoteroUnavailable(f"Zotero web API update failed for {key!r}: {exc}") from exc
    response = getattr(write_client, "request", None)
    headers = getattr(response, "headers", None) or {}
    undo_hint = (
        "To undo, call update_zotero_item again with clean_text=false and fields="
        f"{json.dumps(previous, ensure_ascii=False)}"
    )
    try:
        new_version = int(headers.get("Last-Modified-Version"))
    except (TypeError, ValueError):
        raise ZoteroUnavailable(
            f"Zotero web update of {key!r} was accepted, but the response gave no "
            "Last-Modified-Version, so the local sync cannot be verified. Check the item "
            f"in Zotero before refreshing. {undo_hint}"
        )
    result["version_after"] = new_version

    _wait_for_local_sync(
        key, {c["field"]: c["new"] for c in changes}, new_version, result, sync_timeout_seconds,
        poll_seconds, f"Zotero web update of {key!r} succeeded (version {new_version}), but", f" {undo_hint}",
    )
    return result


def _wait_for_local_sync(
    key: str,
    expected: dict[str, Any],
    version: int,
    result: dict[str, Any],
    sync_timeout_seconds: float,
    poll_seconds: float,
    lead: str,
    tail: str,
) -> None:
    """Wait until the local Zotero shows `expected` at `version` or later; a timeout is an error."""

    read_client = ZoteroClient()
    start = time.monotonic()
    last_error: Exception | None = None
    while True:
        try:
            if _local_matches(read_client.item(key), expected, version):
                result["synced_locally"] = True
                result["sync_seconds"] = round(time.monotonic() - start, 1)
                return
        except ZoteroUnavailable as exc:
            last_error = exc
        if time.monotonic() - start >= sync_timeout_seconds:
            detail = f" (last local error: {last_error})" if last_error else ""
            raise ZoteroUnavailable(
                f"{lead} the local Zotero had not synced it after {sync_timeout_seconds:.0f} s{detail}. "
                "Word citations render from the local library: sync Zotero (or wait) before "
                f"refreshing.{tail}"
            )
        time.sleep(poll_seconds)
