from __future__ import annotations

import asyncio
import json
from pathlib import Path
from types import SimpleNamespace

import pytest

from research_mcp.benchmark import (
    BenchmarkCase,
    EventLog,
    TokenUsage,
    _guard_write_arguments,
    _run_case,
    _word_process_ids,
    load_cases,
)


class FakeModelClient:
    def __init__(self):
        self.calls = 0
        self.requests = []

    def respond(self, **kwargs):
        self.calls += 1
        self.requests.append(kwargs)
        if self.calls == 1:
            return {
                "id": "resp-1",
                "output": [
                    {
                        "type": "function_call",
                        "call_id": "call-1",
                        "name": "audit_document",
                        "arguments": '{"path":"fixture.docx"}',
                    }
                ],
                "usage": {"input_tokens": 100, "output_tokens": 20, "total_tokens": 120},
            }
        return {
            "id": "resp-2",
            "output": [
                {
                    "type": "message",
                    "content": [{"type": "output_text", "text": "Audit complete."}],
                }
            ],
            "usage": {"input_tokens": 40, "output_tokens": 10, "total_tokens": 50},
        }


class FakeMCPResult:
    is_error = False
    structured_content = {"table_count": 1}
    content = []


class FakeSession:
    def __init__(self):
        self.calls = []

    async def call_tool(self, name, arguments):
        self.calls.append((name, arguments))
        return FakeMCPResult()


def fake_tool(*, read_only=True, name="audit_document"):
    return SimpleNamespace(
        name=name,
        description="Audit.",
        input_schema={
            "type": "object",
            "properties": {"path": {"type": "string"}},
            "required": ["path"],
        },
        annotations=SimpleNamespace(read_only_hint=read_only),
    )


def test_case_catalog_distinguishes_current_and_planned_capabilities():
    cases = load_cases(Path("benchmarks/cases.json"))
    ids = {case.id for case in cases}
    assert len(ids) == len(cases)
    assert {
        "audit_document",
        "list_bookmarks",
        "search_zotero",
        "clean_bookmarks",
        "format_tables",
        "insert_zotero_citation_word",
    } <= ids
    # Only capabilities with no exposed tool stay planned. `edit_table_cells`
    # was promoted once `edit_document` landed: it is `replace_text` at a cell
    # paragraph index, so the tool it used to name will never exist. The
    # self-maintaining form of this check is in tests/test_benchmark_verifiers.py.
    planned = {case.id for case in cases if case.tier == "planned"}
    assert planned == {"batch_zotero_citations"}
    executable = {case.id for case in cases if case.tier != "planned"}
    assert {
        "insert_sentence",
        "insert_paragraph",
        "edit_sentence",
        "insert_table",
        "edit_table_cells",
        "reply_to_editorial_comment",
        "partial_batch_around_citation",
    } <= executable


def test_token_usage_sums_each_agent_turn():
    usage = TokenUsage()
    usage.add({"input_tokens": 10, "output_tokens": 5, "total_tokens": 15})
    usage.add({"input_tokens": 7, "output_tokens": 3, "total_tokens": 10})
    assert (usage.input_tokens, usage.output_tokens, usage.total_tokens) == (17, 8, 25)


def test_write_guard_confines_output_to_run_directory(tmp_path):
    tool = fake_tool(read_only=False)
    allowed = tmp_path / "run" / "out.docx"
    _guard_write_arguments(tool, {"output": str(allowed)}, tmp_path / "run")
    with pytest.raises(RuntimeError, match="inside run directory"):
        _guard_write_arguments(tool, {"output": str(tmp_path / "outside.docx")}, tmp_path / "run")
    with pytest.raises(RuntimeError, match="without a versioned output"):
        _guard_write_arguments(tool, {}, tmp_path / "run")


def test_word_process_inventory_parses_tasklist_without_touching_word(monkeypatch):
    monkeypatch.setattr("research_mcp.benchmark.sys.platform", "win32")
    monkeypatch.setattr(
        "research_mcp.benchmark.subprocess.run",
        lambda *args, **kwargs: SimpleNamespace(
            returncode=0,
            stdout='"WINWORD.EXE","1234","Console","1","10,000 K"\n'
            '"WINWORD.EXE","5678","Console","1","11,000 K"\n',
            stderr="",
        ),
    )
    assert _word_process_ids() == [1234, 5678]


def test_agent_loop_records_server_usage_and_verifies_tool_result(tmp_path):
    case = BenchmarkCase(
        id="audit",
        description="audit",
        tier="read_only",
        required_tools=("audit_document",),
        prompt="Audit fixture.docx",
        verifier={"kind": "tool_result", "tool": "audit_document", "contains": {"table_count": 1}},
    )
    client = FakeModelClient()
    session = FakeSession()
    result = asyncio.run(
        _run_case(
            case=case,
            repeat=1,
            toolset_name="minimal",
            all_tools={"audit_document": fake_tool()},
            session=session,
            client=client,
            variables={},
            run_dir=tmp_path,
            max_turns=3,
            max_output_tokens=256,
            log=EventLog(tmp_path / "events.jsonl"),
        )
    )
    assert result.status == "passed"
    assert result.success
    assert result.usage.total_tokens == 170
    assert result.tokens_to_verification == 120
    assert result.first_verified_turn == 1
    assert result.clean_termination
    assert result.model_turns == 2
    assert result.tool_calls == 1
    assert session.calls == [("audit_document", {"path": "fixture.docx"})]
    assert client.requests[1]["previous_response_id"] is None
    assert client.requests[1]["input_items"][1] == {
        "role": "user",
        "content": "Audit fixture.docx",
    }
    assert client.requests[1]["input_items"][-1]["type"] == "function_call_output"
    assert Path(result.transcript_path).exists()


def test_missing_tool_is_unsupported_without_calling_model(tmp_path):
    case = BenchmarkCase(
        id="future",
        description="future",
        tier="planned",
        required_tools=("edit_document",),
        prompt="Edit.",
        verifier={"kind": "planned"},
    )
    client = FakeModelClient()
    result = asyncio.run(
        _run_case(
            case=case,
            repeat=1,
            toolset_name="all",
            all_tools={},
            session=FakeSession(),
            client=client,
            variables={},
            run_dir=tmp_path,
            max_turns=3,
            max_output_tokens=256,
            log=EventLog(tmp_path / "events.jsonl"),
        )
    )
    assert result.status == "unsupported"
    assert result.missing_tools == ["edit_document"]
    assert client.calls == 0


def test_verified_state_with_repeated_calls_is_reported_as_overrun(tmp_path):
    class RepeatingClient:
        def respond(self, **kwargs):
            return {
                "id": f"resp-{id(kwargs)}",
                "output": [
                    {
                        "type": "function_call",
                        "call_id": f"call-{id(kwargs)}",
                        "name": "audit_document",
                        "arguments": '{"path":"fixture.docx"}',
                    }
                ],
                "usage": {"input_tokens": 10, "output_tokens": 5, "total_tokens": 15},
            }

    case = BenchmarkCase(
        id="audit",
        description="audit",
        tier="read_only",
        required_tools=("audit_document",),
        prompt="Audit fixture.docx",
        verifier={"kind": "tool_result", "tool": "audit_document", "contains": {"table_count": 1}},
    )
    result = asyncio.run(
        _run_case(
            case=case,
            repeat=1,
            toolset_name="all",
            all_tools={"audit_document": fake_tool()},
            session=FakeSession(),
            client=RepeatingClient(),
            variables={},
            run_dir=tmp_path,
            max_turns=2,
            max_output_tokens=256,
            log=EventLog(tmp_path / "events.jsonl"),
        )
    )
    assert result.status == "overrun"
    assert result.success
    assert not result.clean_termination
    assert result.tokens_to_verification == 15
    assert result.usage.total_tokens == 30


def test_mutating_task_stops_at_first_verified_postcondition(tmp_path):
    class MutatingClient:
        calls = 0

        def respond(self, **kwargs):
            self.calls += 1
            return {
                "id": "resp-write",
                "output": [
                    {
                        "type": "function_call",
                        "call_id": "call-write",
                        "name": "audit_document",
                        "arguments": json.dumps({"path": "fixture.docx", "output": str(tmp_path / "out.docx")}),
                    }
                ],
                "usage": {"input_tokens": 10, "output_tokens": 5, "total_tokens": 15},
            }

    case = BenchmarkCase(
        id="write",
        description="write",
        tier="versioned",
        required_tools=("audit_document",),
        prompt="Write safely.",
        verifier={"kind": "tool_result", "tool": "audit_document", "contains": {"table_count": 1}},
    )
    client = MutatingClient()
    session = FakeSession()
    result = asyncio.run(
        _run_case(
            case=case,
            repeat=1,
            toolset_name="minimal",
            all_tools={"audit_document": fake_tool(read_only=False)},
            session=session,
            client=client,
            variables={},
            run_dir=tmp_path,
            max_turns=3,
            max_output_tokens=256,
            log=EventLog(tmp_path / "events.jsonl"),
        )
    )
    assert result.status == "passed_guarded"
    assert result.success
    assert result.stopped_on_verification
    assert result.model_turns == 1
    assert client.calls == 1


def test_malformed_tool_arguments_fail_one_call_not_the_whole_case(tmp_path):
    """A model emitting invalid JSON must not crash the runner.

    `arguments` was only bound inside the try, so a `json.loads` failure hit
    UnboundLocalError on the transcript append and reported the case as a
    harness `error` -- hiding the fact that the model sent a malformed call.
    """

    class MalformedClient:
        calls = 0

        def respond(self, **kwargs):
            self.calls += 1
            if self.calls == 1:
                return {
                    "id": "resp-bad",
                    "output": [
                        {
                            "type": "function_call",
                            "call_id": "c1",
                            "name": "audit_document",
                            "arguments": "{not valid json",
                        }
                    ],
                    "usage": {"input_tokens": 10, "output_tokens": 5, "total_tokens": 15},
                }
            return {
                "id": "resp-done",
                "output": [{"type": "message", "content": [{"type": "output_text", "text": "done"}]}],
                "usage": {"input_tokens": 5, "output_tokens": 2, "total_tokens": 7},
            }

    case = BenchmarkCase(
        id="malformed",
        description="malformed",
        tier="read_only",
        required_tools=("audit_document",),
        prompt="Audit it.",
        verifier={"kind": "tool_result", "tool": "audit_document", "contains": {"table_count": 1}},
    )
    result = asyncio.run(
        _run_case(
            case=case,
            repeat=1,
            toolset_name="minimal",
            all_tools={"audit_document": fake_tool()},
            session=FakeSession(),
            client=MalformedClient(),
            variables={},
            run_dir=tmp_path,
            max_turns=3,
            max_output_tokens=256,
            log=EventLog(tmp_path / "events.jsonl"),
        )
    )
    # The case fails verification, but as a model failure, not a harness crash.
    assert result.status == "failed"
    assert result.error is None
    assert result.tool_calls == 1


class _FakeClient:
    """Stands in for the Unsloth server across the preflight's branches."""

    def __init__(self, *, served: bool, samples: list[str]):
        self.model = "test/model"
        self._served = served
        self._samples = samples
        self.calls: list[str] = []

    def ping(self):
        self.calls.append("ping")
        return {"ok": self._served, "model": self.model, "models": ["other/model"]}

    def load(self):
        self.calls.append("load")
        self._served = True
        return {}

    def unload(self):
        self.calls.append("unload")
        return {}

    def health_probe(self):
        self.calls.append("probe")
        sample = self._samples.pop(0)
        if sample == "__unloaded__":
            raise RuntimeError("Unsloth API HTTP 400: No model loaded.")
        from research_mcp.benchmark import _looks_degenerate

        return {
            "healthy": bool(sample.strip()) and not _looks_degenerate(sample),
            "message": sample,
            "sample": sample,
        }


def _log(tmp_path):
    from research_mcp.benchmark import EventLog

    return EventLog(tmp_path / "events.jsonl")


def test_preflight_loads_a_model_that_is_not_served(tmp_path):
    """Running against an unloaded model wasted a whole run before this."""

    from research_mcp.benchmark import ensure_model_ready

    client = _FakeClient(served=True, samples=["__unloaded__", "ready"])
    ensure_model_ready(client, _log(tmp_path))
    assert client.calls == ["probe", "ping", "load", "probe"]


def test_preflight_leaves_a_healthy_loaded_model_alone(tmp_path):
    from research_mcp.benchmark import ensure_model_ready

    client = _FakeClient(served=True, samples=["ready"])
    ensure_model_ready(client, _log(tmp_path))
    assert client.calls == ["probe"]


def test_preflight_reloads_a_wedged_model_and_continues(tmp_path):
    """A wedged llama-server answers /models and then emits garbage.

    Observed after a tool-schema grammar-compilation failure: 768 tokens of
    "/" characters per case, every case scored as a genuine failure. The old
    ping could not see it because the endpoint responded normally.
    """

    from research_mcp.benchmark import ensure_model_ready

    client = _FakeClient(served=True, samples=["/" * 64, "/" * 64, "ready"])
    ensure_model_ready(client, _log(tmp_path))
    # Loading over a wedged model is a no-op, so the reload path must be reached.
    assert client.calls == ["probe", "ping", "load", "probe", "unload", "load", "probe"]


def test_preflight_refuses_to_benchmark_a_model_that_stays_wedged(tmp_path):
    """Better no run than a run scored against garbage."""

    from research_mcp.benchmark import ensure_model_ready

    client = _FakeClient(served=True, samples=["/" * 64, "/" * 64, "/" * 64])
    with pytest.raises(RuntimeError, match="not usable"):
        ensure_model_ready(client, _log(tmp_path))


def test_degenerate_detector_accepts_real_answers_and_rejects_the_wedge():
    from research_mcp.benchmark import _looks_degenerate

    assert _looks_degenerate("/" * 64)
    assert _looks_degenerate("//// //// //// ////")
    assert not _looks_degenerate("ready")
    assert not _looks_degenerate("I will call inspect_document first.")
    # Short repeats are normal punctuation, not a wedge.
    assert not _looks_degenerate("---")


def test_preflight_reports_a_model_the_server_cannot_serve_at_all(tmp_path):
    """Distinguish "not loaded" from "not available"; only one is fixable here."""

    from research_mcp.benchmark import ensure_model_ready

    client = _FakeClient(served=False, samples=["__unloaded__"])
    with pytest.raises(RuntimeError, match="not available"):
        ensure_model_ready(client, _log(tmp_path))
    assert "load" not in client.calls
