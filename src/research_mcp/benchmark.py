"""Agent/MCP benchmark runner for a local OpenAI-compatible model."""

from __future__ import annotations

import argparse
import asyncio
import csv
import hashlib
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import subprocess
import sys
import time
from typing import Any, Iterable
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
import zipfile

from mcp import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client

from . import config
from .audit import audit_docx
from .fixtures import create_table_fixture
from .ooxml import zotero_package_fields
from .paragraphs import read_paragraphs
from .processes import process_identity_matches
from .tables import MergeGroup, TableFormatSpec, verify_saved_journal_neutral
from .word_service import read_worker_state, shutdown_word_worker, worker_state_is_live


DEFAULT_BASE_URL = "http://127.0.0.1:8888/v1"
DEFAULT_MODEL = "unsloth/Qwen3.8-27B-GGUF"
DEFAULT_CASES = Path("benchmarks/cases.json")
DEFAULT_TIERS = ("read_only", "planned")
SYSTEM_PROMPT = """You are evaluating a research-paper MCP server.
Complete exactly the user's requested task with the provided tools.
Use tool results as authoritative evidence. Never claim success if a tool failed.
Never overwrite a source document. Write only to the exact output path supplied.
Make no unrelated changes. When the requested state is verified, answer briefly.
"""


@dataclass(frozen=True)
class BenchmarkCase:
    id: str
    description: str
    tier: str
    required_tools: tuple[str, ...]
    prompt: str
    verifier: dict[str, Any]


@dataclass
class TokenUsage:
    input_tokens: int = 0
    output_tokens: int = 0
    total_tokens: int = 0

    def add(self, raw: dict[str, Any] | None) -> None:
        data = raw or {}
        self.input_tokens += int(data.get("input_tokens") or 0)
        self.output_tokens += int(data.get("output_tokens") or 0)
        self.total_tokens += int(data.get("total_tokens") or 0)


@dataclass
class CaseResult:
    case_id: str
    description: str
    tier: str
    toolset: str
    repeat: int
    status: str
    success: bool
    clean_termination: bool = False
    stopped_on_verification: bool = False
    first_verified_turn: int | None = None
    tokens_to_verification: int | None = None
    missing_tools: list[str] = field(default_factory=list)
    usage: TokenUsage = field(default_factory=TokenUsage)
    model_turns: int = 0
    tool_calls: int = 0
    latency_seconds: float = 0.0
    verification: dict[str, Any] = field(default_factory=dict)
    final_text: str = ""
    error: str | None = None
    transcript_path: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _read_dotenv(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if not path.exists():
        return values
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        name, value = line.split("=", 1)
        values[name.strip()] = value.strip().strip('"').strip("'")
    return values


def _json_request(url: str, body: dict[str, Any] | None, api_key: str, timeout: float) -> dict[str, Any]:
    headers = {"Authorization": f"Bearer {api_key}", "Accept": "application/json"}
    data = None
    method = "GET"
    if body is not None:
        data = json.dumps(body, ensure_ascii=False).encode("utf-8")
        headers["Content-Type"] = "application/json"
        method = "POST"
    request = Request(url, data=data, headers=headers, method=method)
    try:
        with urlopen(request, timeout=timeout) as response:
            return json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", "replace")[:2000]
        raise RuntimeError(f"Unsloth API HTTP {exc.code}: {detail}") from exc
    except (URLError, TimeoutError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"Unsloth API request failed: {exc}") from exc


class UnslothResponsesClient:
    def __init__(self, *, base_url: str, api_key: str, model: str, timeout: float = 180.0):
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key
        self.model = model
        self.timeout = timeout

    def ping(self) -> dict[str, Any]:
        data = _json_request(f"{self.base_url}/models", None, self.api_key, self.timeout)
        models = [str(item.get("id")) for item in data.get("data", []) if isinstance(item, dict)]
        return {"ok": self.model in models, "model": self.model, "models": models}

    def load(self) -> dict[str, Any]:
        """Load the configured model. Loading a GGUF can take minutes."""

        return _json_request(
            f"{self.base_url}/load",
            {"model_path": self.model},
            self.api_key,
            max(self.timeout, 900.0),
        )

    def unload(self) -> dict[str, Any]:
        return _json_request(
            f"{self.base_url}/unload",
            {"model_path": self.model},
            self.api_key,
            max(self.timeout, 300.0),
        )

    def health_probe(self) -> dict[str, Any]:
        """Ask for one trivial completion and look at what comes back.

        A wedged llama-server answers /models normally and then emits
        degenerate text -- an observed failure produced 768 tokens of "/"
        characters with no message block, which the old ping could not see.
        """

        data = _json_request(
            f"{self.base_url}/responses",
            {
                "model": self.model,
                "input": [{"role": "user", "content": "Reply with exactly: ready"}],
                "max_output_tokens": 64,
                "temperature": 0,
            },
            self.api_key,
            self.timeout,
        )
        message = ""
        reasoning = ""
        for block in data.get("output", []) or []:
            for piece in block.get("content", []) or []:
                if piece.get("type") == "output_text":
                    message += str(piece.get("text") or "")
                elif piece.get("type") == "reasoning_text":
                    reasoning += str(piece.get("reasoning_text") or piece.get("text") or "")
        sample = (message or reasoning).strip()
        healthy = bool(message.strip()) and not _looks_degenerate(sample)
        return {"healthy": healthy, "message": message.strip(), "sample": sample[:200]}

    def respond(
        self,
        *,
        input_items: Any,
        tools: list[dict[str, Any]],
        max_output_tokens: int,
        previous_response_id: str | None = None,
    ) -> dict[str, Any]:
        body: dict[str, Any] = {
            "model": self.model,
            "input": input_items,
            "tools": tools,
            "tool_choice": "auto",
            "temperature": 0,
            "max_output_tokens": max_output_tokens,
        }
        if previous_response_id:
            body["previous_response_id"] = previous_response_id
        return _json_request(f"{self.base_url}/responses", body, self.api_key, self.timeout)


def _looks_degenerate(sample: str) -> bool:
    """True when output is one character repeated, the observed wedge signature."""

    stripped = "".join(sample.split())
    return len(stripped) >= 16 and len(set(stripped)) == 1


def _safe_probe(client: "UnslothResponsesClient") -> dict[str, Any]:
    """Health probe that reports transport failures instead of raising.

    `/models` lists models the server *can* serve, not the one it has loaded,
    so an unloaded model pings fine and then fails the probe with HTTP 400
    "No model loaded". That has to read as "load it", not as a fatal error.
    """

    try:
        return client.health_probe()
    except Exception as exc:  # noqa: BLE001 - any transport failure means unhealthy
        return {"healthy": False, "message": "", "sample": f"{type(exc).__name__}: {exc}"[:300]}


def ensure_model_ready(client: "UnslothResponsesClient", log: "EventLog") -> None:
    """Make the model loaded and answering sanely before any case runs.

    Two failures motivated this, both of which a bare `/models` ping passed:
    the model not being loaded at all, and a llama-server left wedged by a
    tool-schema grammar-compilation failure. Wedged, it answers `/models`
    normally and then emits garbage -- an observed run produced 768 tokens of
    "/" characters per case, and every case was scored as a real failure.
    """

    probe = _safe_probe(client)
    log.write("model_health", model=client.model, healthy=probe["healthy"], sample=probe["sample"])
    if probe["healthy"]:
        return

    available = client.ping()
    if not available["ok"]:
        raise RuntimeError(
            f"model {client.model!r} is not available on this server; "
            f"served models: {available['models']}"
        )

    log.write("model_load_started", model=client.model, reason=probe["sample"])
    client.load()
    probe = _safe_probe(client)
    log.write("model_health", model=client.model, healthy=probe["healthy"],
              sample=probe["sample"], after_load=True)
    if probe["healthy"]:
        return

    # Loading over a wedged model is a no-op; only a full unload clears it.
    log.write("model_reload_started", model=client.model, reason=probe["sample"])
    client.unload()
    client.load()
    probe = _safe_probe(client)
    log.write("model_health", model=client.model, healthy=probe["healthy"],
              sample=probe["sample"], after_reload=True)
    if not probe["healthy"]:
        raise RuntimeError(
            f"model {client.model!r} is not usable after a load and a full reload; "
            f"refusing to benchmark against it. Last probe: {probe['sample']!r}"
        )


class EventLog:
    def __init__(self, path: Path):
        self.path = path
        path.parent.mkdir(parents=True, exist_ok=True)

    def write(self, event: str, **data: Any) -> None:
        record = {"timestamp": _utc_now(), "event": event, **data}
        with self.path.open("a", encoding="utf-8") as stream:
            stream.write(json.dumps(record, ensure_ascii=False, default=str) + "\n")


def load_cases(path: Path) -> tuple[BenchmarkCase, ...]:
    raw = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(raw, list):
        raise ValueError("benchmark case file must contain a JSON list")
    cases: list[BenchmarkCase] = []
    ids: set[str] = set()
    for item in raw:
        case = BenchmarkCase(
            id=str(item["id"]),
            description=str(item["description"]),
            tier=str(item["tier"]),
            required_tools=tuple(str(name) for name in item["required_tools"]),
            prompt=str(item["prompt"]),
            verifier=dict(item.get("verifier") or {}),
        )
        if case.id in ids:
            raise ValueError(f"duplicate benchmark case id: {case.id}")
        if not case.required_tools:
            raise ValueError(f"benchmark case {case.id} has no required tools")
        ids.add(case.id)
        cases.append(case)
    return tuple(cases)


def _render(value: Any, variables: dict[str, str]) -> Any:
    if isinstance(value, str):
        return value.format_map(variables)
    if isinstance(value, list):
        return [_render(item, variables) for item in value]
    if isinstance(value, dict):
        return {key: _render(item, variables) for key, item in value.items()}
    return value


def _mcp_tool_schema(tool: Any) -> dict[str, Any]:
    return {
        "type": "function",
        "name": tool.name,
        "description": tool.description or "",
        "parameters": tool.input_schema,
        "strict": False,
    }


def _content_text(content: Iterable[Any]) -> str:
    return "\n".join(
        str(block.text)
        for block in content
        if getattr(block, "type", None) == "text" and getattr(block, "text", None) is not None
    )


def _tool_result_payload(result: Any) -> dict[str, Any]:
    if result.is_error:
        return {"ok": False, "error": _content_text(result.content)}
    structured = result.structured_content
    if structured is not None:
        return {"ok": True, "result": structured}
    return {"ok": True, "result": _content_text(result.content)}


def _inside(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def _active_word_documents() -> list[dict[str, Any]]:
    """Best-effort read-only preflight of the Word application in the ROT."""

    if sys.platform != "win32":
        return []
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            return []
        documents: list[dict[str, Any]] = []
        for index in range(1, int(word.Documents.Count) + 1):
            document = word.Documents(index)
            documents.append(
                {
                    "name": str(document.Name),
                    "path": str(document.FullName),
                    "saved": bool(document.Saved),
                    "read_only": bool(document.ReadOnly),
                }
            )
        return documents
    finally:
        pythoncom.CoUninitialize()


def _word_process_ids() -> list[int]:
    """Enumerate Word processes without attaching to or changing Word."""

    if sys.platform != "win32":
        return []
    completed = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
        timeout=15,
        check=False,
    )
    if completed.returncode != 0:
        raise RuntimeError(f"could not enumerate Word processes: {completed.stderr.strip()}")
    pids: list[int] = []
    for row in csv.reader(completed.stdout.splitlines()):
        if len(row) < 2 or row[0].lower() != "winword.exe":
            continue
        try:
            pids.append(int(row[1]))
        except ValueError:
            continue
    return sorted(pids)


def _guard_write_arguments(tool: Any, arguments: dict[str, Any], run_dir: Path) -> None:
    annotations = getattr(tool, "annotations", None)
    if annotations is not None and getattr(annotations, "read_only_hint", None) is True:
        return
    if "output" not in arguments:
        raise RuntimeError(
            f"benchmark safety guard refused mutating tool {tool.name!r} without a versioned output path"
        )
    output = Path(str(arguments["output"]))
    if not _inside(output, run_dir):
        raise RuntimeError(f"benchmark output must stay inside run directory: {output}")


def _message_text(response: dict[str, Any]) -> str:
    parts: list[str] = []
    for item in response.get("output", []):
        if item.get("type") != "message":
            continue
        for content in item.get("content", []):
            if content.get("type") == "output_text":
                parts.append(str(content.get("text") or ""))
    return "\n".join(parts).strip()


def _function_calls(response: dict[str, Any]) -> list[dict[str, Any]]:
    return [item for item in response.get("output", []) if item.get("type") == "function_call"]


def _last_successful_result(transcript: list[dict[str, Any]], tool_name: str) -> Any | None:
    for entry in reversed(transcript):
        if entry.get("kind") == "tool" and entry.get("name") == tool_name:
            payload = entry.get("output") or {}
            if payload.get("ok"):
                return payload.get("result")
    return None


def _contains(value: Any, expected: Any) -> bool:
    if isinstance(expected, dict):
        return isinstance(value, dict) and all(key in value and _contains(value[key], item) for key, item in expected.items())
    if isinstance(expected, list):
        return isinstance(value, list) and all(any(_contains(candidate, item) for candidate in value) for item in expected)
    return value == expected


def _deep_has_string(value: Any, expected: str) -> bool:
    if isinstance(value, str):
        return value == expected
    if isinstance(value, dict):
        return any(_deep_has_string(item, expected) for item in value.values())
    if isinstance(value, list):
        return any(_deep_has_string(item, expected) for item in value)
    return False


def verify_case(
    case: BenchmarkCase,
    verifier: dict[str, Any],
    transcript: list[dict[str, Any]],
) -> tuple[bool, dict[str, Any]]:
    kind = verifier.get("kind")
    tool_name = str(verifier.get("tool") or case.required_tools[0])
    tool_result = _last_successful_result(transcript, tool_name)
    if tool_result is None:
        return False, {"reason": f"no successful {tool_name} call"}

    if kind == "tool_result":
        expected = verifier.get("contains", {})
        strings = [str(value) for value in verifier.get("strings", [])]
        checks = {
            "structured_subset": _contains(tool_result, expected),
            "required_strings": all(_deep_has_string(tool_result, value) for value in strings),
        }
        return all(checks.values()), {"checks": checks}

    output = Path(str(verifier.get("output", ""))).resolve()
    if not output.exists():
        return False, {"reason": f"expected output does not exist: {output}"}

    if kind == "bookmark_cleanup":
        source = Path(str(verifier["source"]))
        before = audit_docx(source)
        after = audit_docx(output)
        absent = [str(name) for name in verifier.get("absent", [])]
        present = [str(name) for name in verifier.get("present", [])]
        checks = {
            "source_hash_unchanged": audit_docx(source).sha256
            == str(verifier.get("_source_sha256_before") or before.sha256),
            "body_text_unchanged": after.body_text_sha256 == before.body_text_sha256,
            "absent": all(name not in after.bookmark_names for name in absent),
            "present": all(name in after.bookmark_names for name in present),
        }
        return all(checks.values()), {"checks": checks, "output_sha256": after.sha256}

    if kind == "table_format":
        source = Path(str(verifier["source"]))
        before = audit_docx(source)
        after = audit_docx(output)
        specs = tuple(
            TableFormatSpec(
                index=int(raw["index"]),
                descriptive_columns=tuple(int(value) for value in raw["descriptive_columns"]),
                merges=tuple(MergeGroup(**merge) for merge in raw.get("merges", [])),
            )
            for raw in verifier["table_specs"]
        )
        formatting_error = None
        try:
            verify_saved_journal_neutral(output, specs)
        except Exception as exc:  # verifier reports the evidence rather than crashing the run
            formatting_error = str(exc)
        expected_merges = tuple(
            (group.start_row, group.end_row, group.column, group.label)
            for spec in specs
            for group in spec.merges
        )
        actual_merges = tuple(merge for table in after.tables for merge in table.vertical_merges)
        checks = {
            "source_hash_unchanged": audit_docx(source).sha256
            == str(verifier.get("_source_sha256_before") or before.sha256),
            # `body_text` includes table cell paragraphs, and a vertical merge
            # physically empties the cells it merges away -- so this check
            # failed every correct merge. `prose_text_sha256` covers top-level
            # prose only and is immune to it, while `cell_text_unchanged`
            # already proves the logical table content is identical.
            "prose_text_unchanged": after.prose_text_sha256 == before.prose_text_sha256,
            "cell_text_unchanged": tuple(x for t in after.tables for x in t.cell_hashes)
            == tuple(x for t in before.tables for x in t.cell_hashes),
            "merges_exact": actual_merges == expected_merges,
            "journal_neutral": formatting_error is None,
        }
        return all(checks.values()), {"checks": checks, "formatting_error": formatting_error}

    if kind == "document_edit":
        source = Path(str(verifier["source"]))
        before = audit_docx(source)
        after = audit_docx(output)
        visible_text = "\n".join(item.text for item in read_paragraphs(output))
        expected_text = [str(value) for value in verifier.get("expected_text", [])]
        absent_text = [str(value) for value in verifier.get("absent_text", [])]
        expected_source_sha = str(verifier.get("_source_sha256_before") or before.sha256)
        checks = {
            "tool_reported_publication": isinstance(tool_result, dict)
            and tool_result.get("status") in {"complete", "partial"}
            and str(tool_result.get("output_sha256", "")).upper() == after.sha256,
            "source_hash_unchanged": audit_docx(source).sha256 == expected_source_sha,
            "expected_text": all(value in visible_text for value in expected_text),
            "absent_text": all(value not in visible_text for value in absent_text),
            "table_delta": after.table_count
            == before.table_count + int(verifier.get("table_delta", 0)),
            "fields_preserved": after.field_count == before.field_count,
            "drawings_preserved": after.drawing_count == before.drawing_count,
            "media_preserved": after.live_media_count == before.live_media_count,
        }
        # A reply adds a comment, so strict equality is only right when the
        # case does not declare a delta. Every other case keeps the guarantee.
        checks["comments_preserved"] = after.live_comment_count == (
            before.live_comment_count + int(verifier.get("comment_delta", 0))
        )
        if "paragraph_delta" in verifier:
            checks["paragraph_delta"] = after.paragraph_count == (
                before.paragraph_count + int(verifier["paragraph_delta"])
            )
        # "partial" without checking which operations applied is narration.
        # A batch that silently published a refused operation would otherwise
        # pass on expected_text alone.
        expected_statuses = verifier.get("operation_statuses")
        if expected_statuses is not None:
            reported = [
                str(item.get("status"))
                for item in (tool_result or {}).get("operations", [])
                if isinstance(item, dict)
            ]
            checks["operation_statuses"] = reported == [
                str(value) for value in expected_statuses
            ]
        new_table = verifier.get("new_table")
        if new_table is not None:
            expected_hashes = tuple(
                hashlib.sha256(" ".join(str(value).split()).encode("utf-8")).hexdigest()
                for row in new_table
                for value in row
            )
            inserted = [table for table in after.tables if table.cell_hashes == expected_hashes]
            before_identities = [
                (table.rows, table.columns, table.cell_hashes, table.vertical_merges)
                for table in before.tables
            ]
            remaining = list(after.tables)
            if len(inserted) == 1:
                remaining.remove(inserted[0])
            checks["new_table_exact"] = len(inserted) == 1
            checks["existing_tables_preserved"] = [
                (table.rows, table.columns, table.cell_hashes, table.vertical_merges)
                for table in remaining
            ] == before_identities
        return all(checks.values()), {"checks": checks, "output_sha256": after.sha256}

    if kind == "zotero_citation":
        expected_key = str(verifier["item_key"])
        audit = audit_docx(output)
        fields = zotero_package_fields(output)
        with zipfile.ZipFile(output) as archive:
            raw = archive.read("word/document.xml")
        rendered = [result for _, result in fields.item_results]
        checks = {
            "one_item_field": audit.zotero_item_fields == 1,
            "expected_item_key": any(
                expected_key in keys for keys in fields.item_keys
            ),
            "rendered": bool(rendered) and all(value and "{Citation}" not in value for value in rendered),
        }
        return all(checks.values()), {"checks": checks, "rendered": rendered}

    return False, {"reason": f"unknown verifier kind: {kind!r}"}


async def _run_case(
    *,
    case: BenchmarkCase,
    repeat: int,
    toolset_name: str,
    all_tools: dict[str, Any],
    session: ClientSession,
    client: UnslothResponsesClient,
    variables: dict[str, str],
    run_dir: Path,
    max_turns: int,
    max_output_tokens: int,
    log: EventLog,
) -> CaseResult:
    started = time.monotonic()
    missing = [name for name in case.required_tools if name not in all_tools]
    result = CaseResult(
        case_id=case.id,
        description=case.description,
        tier=case.tier,
        toolset=toolset_name,
        repeat=repeat,
        status="unsupported" if missing else "running",
        success=False,
        missing_tools=missing,
    )
    if missing:
        result.error = f"MCP does not expose required tools: {', '.join(missing)}"
        result.latency_seconds = time.monotonic() - started
        log.write("case_end", result=result.to_dict())
        return result

    exposed = all_tools if toolset_name == "all" else {name: all_tools[name] for name in case.required_tools}
    required_mutation = any(
        not (
            getattr(all_tools[name], "annotations", None) is not None
            and getattr(all_tools[name].annotations, "read_only_hint", None) is True
        )
        for name in case.required_tools
    )
    openai_tools = [_mcp_tool_schema(tool) for tool in exposed.values()]
    prompt = str(_render(case.prompt, variables))
    verifier = dict(_render(case.verifier, variables))
    if "source" in verifier:
        verifier["_source_sha256_before"] = audit_docx(verifier["source"]).sha256
    transcript: list[dict[str, Any]] = []
    conversation_input: list[dict[str, Any]] = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": prompt},
    ]
    log.write("case_start", case=case.id, toolset=toolset_name, repeat=repeat, prompt=prompt)

    try:
        latest_verification: dict[str, Any] = {}
        for turn in range(1, max_turns + 1):
            response = await asyncio.to_thread(
                client.respond,
                input_items=conversation_input,
                tools=openai_tools,
                max_output_tokens=max_output_tokens,
                # llama-server's response-id cache is process-global and can
                # replay unrelated prior context. Send the complete case history
                # explicitly so each benchmark case is genuinely isolated.
                previous_response_id=None,
            )
            result.model_turns += 1
            result.usage.add(response.get("usage"))
            transcript.append({"kind": "model", "turn": turn, "response": response})
            log.write(
                "model_response",
                case=case.id,
                toolset=toolset_name,
                repeat=repeat,
                turn=turn,
                response_id=response.get("id"),
                usage=response.get("usage"),
                output=response.get("output"),
            )
            calls = _function_calls(response)
            if not calls:
                result.final_text = _message_text(response)
                result.clean_termination = True
                break

            outputs: list[dict[str, Any]] = []
            for call in calls:
                result.tool_calls += 1
                name = str(call.get("name") or "")
                call_id = str(call.get("call_id") or "")
                # Assigned up front: if `json.loads` raises, the transcript
                # append below used to hit UnboundLocalError and kill the whole
                # case with a runner error, hiding the model's malformed call.
                arguments: dict[str, Any] = {}
                if name not in exposed:
                    payload = {"ok": False, "error": f"tool not exposed in this benchmark run: {name}"}
                else:
                    try:
                        arguments = json.loads(call.get("arguments") or "{}")
                        if not isinstance(arguments, dict):
                            raise ValueError("tool arguments must be a JSON object")
                        _guard_write_arguments(exposed[name], arguments, run_dir)
                        log.write(
                            "tool_call",
                            case=case.id,
                            toolset=toolset_name,
                            repeat=repeat,
                            name=name,
                            arguments=arguments,
                        )
                        mcp_result = await session.call_tool(name, arguments)
                        payload = _tool_result_payload(mcp_result)
                    except Exception as exc:
                        payload = {"ok": False, "error": str(exc)}
                transcript.append(
                    {"kind": "tool", "name": name, "call_id": call_id, "arguments": arguments, "output": payload}
                )
                log.write(
                    "tool_result",
                    case=case.id,
                    toolset=toolset_name,
                    repeat=repeat,
                    name=name,
                    output=payload,
                )
                outputs.append(
                    {
                        "type": "function_call_output",
                        "call_id": call_id,
                        "output": json.dumps(payload, ensure_ascii=False, default=str),
                    }
                )
            conversation_input.extend(
                {
                    "type": "function_call",
                    "call_id": str(call.get("call_id") or ""),
                    "name": str(call.get("name") or ""),
                    "arguments": str(call.get("arguments") or "{}"),
                }
                for call in calls
            )
            conversation_input.extend(outputs)
            verified, latest_verification = verify_case(case, verifier, transcript)
            if verified and result.first_verified_turn is None:
                result.success = True
                result.first_verified_turn = turn
                result.tokens_to_verification = result.usage.total_tokens
                result.verification = latest_verification
                log.write(
                    "state_verified",
                    case=case.id,
                    toolset=toolset_name,
                    repeat=repeat,
                    turn=turn,
                    total_tokens=result.usage.total_tokens,
                    verification=latest_verification,
                )
                if required_mutation:
                    result.status = "passed_guarded"
                    result.stopped_on_verification = True
                    log.write(
                        "guarded_stop",
                        case=case.id,
                        toolset=toolset_name,
                        repeat=repeat,
                        reason="mutating task reached its independently verified postcondition",
                    )
                    break
        else:
            if result.success:
                result.status = "overrun"
                result.error = f"requested state verified, but model exceeded {max_turns} turns without stopping"
            else:
                raise RuntimeError(f"model exceeded {max_turns} turns")

        if result.status not in {"overrun", "passed_guarded"}:
            success, verification = verify_case(case, verifier, transcript)
            result.success = success
            result.status = "passed" if success else "failed"
            result.verification = verification
            if success and result.first_verified_turn is None:
                result.first_verified_turn = result.model_turns
                result.tokens_to_verification = result.usage.total_tokens
    except Exception as exc:
        result.status = "error"
        result.error = str(exc)
    finally:
        result.latency_seconds = time.monotonic() - started
        transcript_dir = run_dir / "transcripts"
        transcript_dir.mkdir(parents=True, exist_ok=True)
        transcript_path = transcript_dir / f"{case.id}--{toolset_name}--r{repeat}.json"
        transcript_path.write_text(json.dumps(transcript, indent=2, ensure_ascii=False, default=str), encoding="utf-8")
        result.transcript_path = str(transcript_path)
        log.write("case_end", result=result.to_dict())
    return result


def _write_outputs(run_dir: Path, metadata: dict[str, Any], results: list[CaseResult]) -> None:
    summary = {**metadata, "updated_at": _utc_now(), "results": [item.to_dict() for item in results]}
    (run_dir / "summary.json").write_text(
        json.dumps(summary, indent=2, ensure_ascii=False, default=str), encoding="utf-8"
    )

    lines = [
        "# Research MCP Agent Benchmark",
        "",
        f"- Model: `{metadata['model']}`",
        f"- MCP tools discovered: {', '.join(f'`{name}`' for name in metadata['mcp_tools'])}",
        f"- Started: {metadata['started_at']}",
        "",
        "| Case | Tier | Toolset | Status | Tokens (in/out/total) | Tokens to verify | Turns | Calls | Seconds |",
        "|---|---|---|---|---:|---:|---:|---:|---:|",
    ]
    for item in results:
        lines.append(
            f"| `{item.case_id}` | {item.tier} | {item.toolset} | {item.status} | "
            f"{item.usage.input_tokens}/{item.usage.output_tokens}/{item.usage.total_tokens} | "
            f"{item.tokens_to_verification if item.tokens_to_verification is not None else '-'} | "
            f"{item.model_turns} | {item.tool_calls} | {item.latency_seconds:.2f} |"
        )

    completed = [item for item in results if item.status not in {"unsupported", "skipped"}]
    passed = [item for item in completed if item.success]
    clean = [item for item in passed if item.clean_termination]
    unsupported = [item for item in results if item.status == "unsupported"]
    total_tokens = sum(item.usage.total_tokens for item in completed)
    lines.extend(
        [
            "",
            "## Totals",
            "",
            f"- Requested state independently verified: {len(passed)}/{len(completed)} executed runs.",
            f"- Verified and cleanly terminated: {len(clean)}/{len(completed)} executed runs.",
            f"- Safely stopped at a mutating postcondition: {sum(item.stopped_on_verification for item in passed)}.",
            f"- Unsupported capability runs: {len(unsupported)}.",
            f"- Total model tokens: {total_tokens}.",
            f"- Tokens per verified success: {total_tokens / len(passed):.1f}." if passed else "- Tokens per verified success: n/a.",
            f"- Tokens through first verification: {sum(item.tokens_to_verification or 0 for item in passed)}.",
            "",
            "## All-tools versus minimal-tools tradeoff",
            "",
            "Positive token delta means exposing the full MCP surface cost more tokens.",
            "",
            "| Case | All status/tokens | Minimal status/tokens | Token delta (all-minimal) |",
            "|---|---:|---:|---:|",
        ]
    )
    keys = sorted({(item.case_id, item.repeat) for item in results})
    for case_id, repeat in keys:
        variants = {
            item.toolset: item
            for item in results
            if item.case_id == case_id and item.repeat == repeat and item.toolset in {"all", "minimal"}
        }
        if set(variants) != {"all", "minimal"}:
            continue
        all_item, minimal_item = variants["all"], variants["minimal"]
        delta = all_item.usage.total_tokens - minimal_item.usage.total_tokens
        lines.append(
            f"| `{case_id}` r{repeat} | {all_item.status}/{all_item.usage.total_tokens} | "
            f"{minimal_item.status}/{minimal_item.usage.total_tokens} | {delta:+d} |"
        )

    if unsupported:
        lines.extend(["", "## Capability gaps", ""])
        seen: set[tuple[str, tuple[str, ...]]] = set()
        for item in unsupported:
            key = (item.case_id, tuple(item.missing_tools))
            if key in seen:
                continue
            seen.add(key)
            lines.append(f"- `{item.case_id}`: missing {', '.join(f'`{name}`' for name in item.missing_tools)}")

    (run_dir / "report.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


async def run_benchmark(args: argparse.Namespace) -> Path:
    workspace = Path(args.workspace).resolve()
    case_path = Path(args.cases)
    if not case_path.is_absolute():
        case_path = workspace / case_path
    cases = load_cases(case_path)
    selected_ids = set(args.case or [])
    selected_tiers = set(args.tiers)
    selected = [
        case for case in cases if case.tier in selected_tiers and (not selected_ids or case.id in selected_ids)
    ]
    if selected_ids - {case.id for case in cases}:
        raise ValueError(f"unknown cases: {sorted(selected_ids - {case.id for case in cases})}")
    if not selected:
        raise ValueError("no benchmark cases selected")

    run_dir = Path(args.output_dir).resolve() if args.output_dir else workspace / "benchmark_runs" / datetime.now().strftime("%Y%m%d-%H%M%S")
    if run_dir.exists() and any(run_dir.iterdir()):
        raise ValueError(f"benchmark output directory must be new or empty: {run_dir}")
    run_dir.mkdir(parents=True, exist_ok=True)
    log = EventLog(run_dir / "events.jsonl")

    dotenv = _read_dotenv(workspace / ".env")
    api_key = os.environ.get("UNSLOTH_API_KEY") or dotenv.get("UNSLOTH_API_KEY")
    if not api_key:
        raise RuntimeError("UNSLOTH_API_KEY is missing from the environment and .env")
    client = UnslothResponsesClient(
        base_url=args.base_url,
        api_key=api_key,
        model=args.model,
        timeout=args.api_timeout,
    )
    await asyncio.to_thread(ensure_model_ready, client, log)

    worker_before = read_worker_state()
    worker_before_live = worker_state_is_live(worker_before)
    worker_before_identity = (
        (int(worker_before["pid"]), int(worker_before["process_created"]))
        if worker_before_live and worker_before is not None
        else None
    )
    if any(case.tier == "zotero_live" for case in selected):
        open_documents = await asyncio.to_thread(_active_word_documents)
        word_pids = await asyncio.to_thread(_word_process_ids)
        worker_word_pid = int(worker_before["word_pid"]) if worker_before_live and worker_before else None
        unrelated_word_pids = [pid for pid in word_pids if pid != worker_word_pid]
        log.write(
            "word_preflight",
            documents=open_documents,
            word_pids=word_pids,
            project_worker_word_pid=worker_word_pid,
            unrelated_word_pids=unrelated_word_pids,
        )
        if (open_documents or unrelated_word_pids) and not args.allow_open_word_documents:
            names = [item["path"] for item in open_documents]
            raise RuntimeError(
                "refusing live Zotero benchmark while a non-worker Word process or open document exists; "
                "close unrelated Word work or explicitly pass --allow-open-word-documents. "
                f"documents={names}, unrelated_word_pids={unrelated_word_pids}"
            )

    fixture = run_dir / "word_fixture.docx"
    if any("{fixture}" in json.dumps({"prompt": case.prompt, "verifier": case.verifier}) for case in selected):
        await asyncio.to_thread(create_table_fixture, fixture)
        log.write("fixture_created", path=str(fixture), audit=audit_docx(fixture).to_dict())

    parameters = StdioServerParameters(command=sys.executable, args=["-m", "research_mcp.server"])
    results: list[CaseResult] = []
    metadata: dict[str, Any] = {
        "started_at": _utc_now(),
        "workspace": str(workspace),
        "run_dir": str(run_dir),
        "model": args.model,
        "base_url": args.base_url,
        "tiers": sorted(selected_tiers),
        "toolset_mode": args.toolset,
        "mcp_tools": [],
    }

    async with stdio_client(parameters) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            listed = await session.list_tools()
            all_tools = {tool.name: tool for tool in listed.tools}
            metadata["mcp_tools"] = sorted(all_tools)
            log.write("mcp_tools", tools=metadata["mcp_tools"])
            toolsets = ("all", "minimal") if args.toolset == "both" else (args.toolset,)
            for repeat in range(1, args.repeat + 1):
                for case in selected:
                    for toolset_name in toolsets:
                        variant = f"{toolset_name}-r{repeat}"
                        variables = {
                            "workspace": str(workspace),
                            "run_dir": str(run_dir),
                            "fixture": str(fixture),
                            "output": str(run_dir / f"{case.id}--{variant}.docx"),
                            "citation_source": str(workspace / "tests/data/zotero_integration_source.docx"),
                            # A comment case needs a document Word can open
                            # that actually carries threads. The synthetic
                            # OOXML comment fixture the unit tests use is too
                            # bare for Word; point RESEARCH_MCP_BENCH_COMMENT_SOURCE
                            # at a real manuscript with comment threads.
                            "comment_source": config.get("RESEARCH_MCP_BENCH_COMMENT_SOURCE")
                            or str(workspace / "artifacts/commented_manuscript.docx"),
                            "item_donor": str(workspace / "tests/data/zotero_item_donor.docx"),
                            "bibliography_donor": str(workspace / "tests/data/zotero_bibliography_donor.docx"),
                            "variant": variant,
                        }
                        item = await _run_case(
                            case=case,
                            repeat=repeat,
                            toolset_name=toolset_name,
                            all_tools=all_tools,
                            session=session,
                            client=client,
                            variables=variables,
                            run_dir=run_dir,
                            max_turns=args.max_turns,
                            max_output_tokens=args.max_output_tokens,
                            log=log,
                        )
                        results.append(item)
                        _write_outputs(run_dir, metadata, results)
            if any(case.tier == "zotero_live" for case in selected):
                worker_after = read_worker_state()
                worker_after_live = worker_state_is_live(worker_after)
                worker_after_identity = (
                    (int(worker_after["pid"]), int(worker_after["process_created"]))
                    if worker_after_live and worker_after is not None
                    else None
                )
                if worker_after_live and worker_after_identity != worker_before_identity:
                    word_pid = int(worker_after["word_pid"])
                    word_created = int(worker_after["word_process_created"])
                    cleanup = await asyncio.to_thread(shutdown_word_worker)
                    deadline = time.monotonic() + 20.0
                    while process_identity_matches(word_pid, word_created) and time.monotonic() < deadline:
                        await asyncio.sleep(0.2)
                    word_exited = not process_identity_matches(word_pid, word_created)
                    metadata["worker_cleanup"] = {
                        "ownership": "created_during_benchmark",
                        "response": cleanup,
                        "word_pid": word_pid,
                        "word_exited": word_exited,
                    }
                    log.write("worker_cleanup", **metadata["worker_cleanup"])
                    if not word_exited:
                        raise RuntimeError(
                            f"benchmark-owned Word process {word_pid} did not exit after graceful worker shutdown"
                        )
                elif worker_after_live:
                    metadata["worker_cleanup"] = {"ownership": "preexisting", "action": "left_running"}
                    log.write("worker_cleanup", **metadata["worker_cleanup"])
                else:
                    metadata["worker_cleanup"] = {"ownership": "none_live_after_run", "action": "none"}
                    log.write("worker_cleanup", **metadata["worker_cleanup"])
    _write_outputs(run_dir, metadata, results)
    return run_dir


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Benchmark an Unsloth agent against the research MCP")
    parser.add_argument("--workspace", default=str(Path.cwd()))
    parser.add_argument("--cases", default=str(DEFAULT_CASES))
    parser.add_argument("--output-dir")
    parser.add_argument("--base-url", default=os.environ.get("UNSLOTH_BASE_URL", DEFAULT_BASE_URL))
    parser.add_argument("--model", default=os.environ.get("UNSLOTH_MODEL", DEFAULT_MODEL))
    parser.add_argument("--tiers", nargs="+", default=list(DEFAULT_TIERS))
    parser.add_argument("--case", action="append", help="run only this case id; repeatable")
    parser.add_argument("--toolset", choices=("all", "minimal", "both"), default="both")
    parser.add_argument("--repeat", type=int, default=1)
    parser.add_argument("--max-turns", type=int, default=6)
    parser.add_argument("--max-output-tokens", type=int, default=768)
    parser.add_argument("--api-timeout", type=float, default=180.0)
    parser.add_argument(
        "--allow-open-word-documents",
        action="store_true",
        help="explicitly accept live-integration risk while another Word document is open",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.repeat < 1 or args.max_turns < 1 or args.max_output_tokens < 64:
        raise SystemExit("repeat/max-turns must be positive and max-output-tokens must be at least 64")
    run_dir = asyncio.run(run_benchmark(args))
    print(run_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
