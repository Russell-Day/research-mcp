"""Guards for the two cost axes the agent benchmark does not currently report.

The runner answers "did the model succeed and what did the run cost". These
tests pin the two questions that decide whether an MCP surface is *usable*:

1. what attaching the server costs on every turn, before any work happens; and
2. what one tool call adds to the context, permanently.

Everything here is offline: no model, no Word, no Zotero, no integration lease,
so this file is safe to run while another agent is benchmarking.
"""

from __future__ import annotations

import asyncio
import inspect
import json
from pathlib import Path

import pytest

from research_mcp.server import server
from research_mcp.benchmark_metrics import (
    InteractionQuality,
    SchemaCost,
    calibrate_chars_per_token,
    fit_payload_chars_per_token,
    interaction_quality,
    json_chars,
    load_events,
    payload_cost,
    predict_schema_tokens,
    schema_context_cost,
)


# Copies of recorded runs live under tests/data. `benchmark_runs/` is a scratch
# namespace that gains a new timestamped directory on every invocation and will
# eventually be cleaned; pointing these tests at it would let the measurement
# layer and both defect reproductions go silently green-by-skip.
RUNS = Path(__file__).resolve().parent / "data" / "benchmark_runs"
# Two independent runs recorded identical turn-1 readings for the same case,
# which is what makes the schema delta a measurement rather than an observation.
REPRODUCED_RUNS = ("read_only", "read_only_repeat")


def _events(name: str) -> list[dict]:
    return load_events(RUNS / name)


def test_recorded_run_fixtures_are_present():
    """Unskipped on purpose: deleting the fixtures must turn the suite red.

    Every empirical test below is `skipif`-guarded, so without this guard a
    missing fixture would look like a clean pass.
    """

    for name in REPRODUCED_RUNS + ("versioned", "zotero_live"):
        assert (RUNS / name / "events.jsonl").exists(), f"missing recorded run {name!r}"


def _server_schema_chars() -> dict[str, int]:
    """Serialized size of each tool schema exactly as the runner sends it."""

    tools = asyncio.run(server.list_tools())
    return {
        tool.name: json_chars(
            {
                "type": "function",
                "name": tool.name,
                "description": inspect.cleandoc(tool.description or ""),
                "parameters": tool.input_schema,
                "strict": False,
            }
        )
        for tool in tools
    }


# --------------------------------------------------------------------------
# Axis 1: context inclusion
# --------------------------------------------------------------------------


def test_schema_delta_isolates_the_tool_block_from_the_prompt():
    """Turn 1 has no history, so the all/minimal difference is schemas only."""

    events = [
        {"event": "model_response", "case": "c", "toolset": "all", "repeat": 1,
         "turn": 1, "usage": {"input_tokens": 1178}},
        {"event": "model_response", "case": "c", "toolset": "minimal", "repeat": 1,
         "turn": 1, "usage": {"input_tokens": 455}},
        # Later turns carry tool results and must never contaminate the delta.
        {"event": "model_response", "case": "c", "toolset": "all", "repeat": 1,
         "turn": 2, "usage": {"input_tokens": 99999}},
    ]
    costs = schema_context_cost(events)
    assert len(costs) == 1
    assert costs[0].withheld_schema_tokens == 723
    assert costs[0].schema_share_of_first_turn == pytest.approx(723 / 1178)


def test_schema_delta_needs_both_conditions_of_the_same_repeat():
    """An unpaired reading is not a measurement and must not be reported."""

    events = [
        {"event": "model_response", "case": "c", "toolset": "all", "repeat": 1,
         "turn": 1, "usage": {"input_tokens": 1178}},
        {"event": "model_response", "case": "c", "toolset": "minimal", "repeat": 2,
         "turn": 1, "usage": {"input_tokens": 455}},
    ]
    assert schema_context_cost(events) == []


@pytest.mark.skipif(
    not all((RUNS / name / "events.jsonl").exists() for name in REPRODUCED_RUNS),
    reason="requires the recorded read-only benchmark runs",
)
def test_schema_context_cost_is_reproducible_across_independent_runs():
    """The same case, run at two different times, must report the same block."""

    measured = {}
    for name in REPRODUCED_RUNS:
        for cost in schema_context_cost(_events(name)):
            measured.setdefault(cost.case_id, set()).add(cost.withheld_schema_tokens)

    assert measured["search_zotero"] == {723}, measured
    # Every read-only case pays over half its first turn for schemas it never
    # uses. That ratio is the headline "context inclusion" number.
    for name in REPRODUCED_RUNS:
        for cost in schema_context_cost(_events(name)):
            assert cost.schema_share_of_first_turn > 0.5, cost


@pytest.mark.skipif(
    not (RUNS / REPRODUCED_RUNS[0] / "events.jsonl").exists(),
    reason="requires the recorded read-only benchmark runs",
)
def test_chars_per_token_calibration_is_stable_enough_to_predict_a_new_tool():
    """Prediction is only legitimate if the ratio is a constant, not a fit.

    Five cases each withheld a different subset of schemas. If the derived
    chars-per-token values agree, one serialized schema predicts the per-turn
    cost of a tool before it is written -- which is the whole point, since
    `inspect_document` and `edit_document` are next on the build order.
    """

    # Calibrate against the schema sizes as they were when the runs were
    # recorded, not against the live server. The surface grows as the build
    # order lands -- `inspect_document` already added ~855 characters -- and
    # calibrating a frozen delta against a moved surface would quietly report
    # a ratio for a measurement that never happened.
    schema_chars = json.loads(
        (RUNS / "schema_chars.json").read_text(encoding="utf-8")
    )
    assert sum(schema_chars.values()) == 3071, schema_chars
    costs = [
        cost
        for name in REPRODUCED_RUNS + ("versioned",)
        if (RUNS / name / "events.jsonl").exists()
        for cost in schema_context_cost(_events(name))
    ]
    ratios = calibrate_chars_per_token(costs, schema_chars)
    assert len(ratios) >= 4, ratios
    spread = max(ratios.values()) - min(ratios.values())
    assert spread < 0.15, ratios  # observed spread is ~0.04 chars/token


def test_recorded_schema_sizes_have_only_documented_drift():
    """The historical map must not drift beyond the documented changes.

    The benchmark fixtures predate the collection-search refinement, the
    citation-anchor contract, and current in-place workflow clarifications.
    Their old sizes remain necessary for calibrating
    the runs they actually recorded, so these differences are explicit
    rather than silently replacing the historical baseline. Any other change
    still fails here and requires a deliberate benchmark update.
    """

    frozen = json.loads((RUNS / "schema_chars.json").read_text(encoding="utf-8"))
    live = _server_schema_chars()
    assert set(frozen) <= set(live), sorted(set(frozen) - set(live))
    drifted = {name: (size, live[name]) for name, size in frozen.items() if live[name] != size}
    assert drifted == {
        "clean_bookmarks": (474, 668),
        "format_tables": (499, 2019),
        "insert_zotero_citation": (1127, 1597),
        "list_bookmarks": (307, 300),
        "search_zotero": (369, 941),
    }, (
        f"unexpected recorded tool schema drift {drifted}; re-record the "
        "benchmark runs in tests/data/benchmark_runs and refresh schema_chars.json"
    )


def test_predicted_schema_tokens_scale_with_schema_size():
    small = {"name": "a", "parameters": {"type": "object", "properties": {}}}
    large = {
        "name": "a",
        "parameters": {
            "type": "object",
            "properties": {f"field_{index}": {"type": "string"} for index in range(20)},
        },
    }
    assert predict_schema_tokens(small) < predict_schema_tokens(large)
    assert predict_schema_tokens(large) > 0


# --------------------------------------------------------------------------
# Axis 2: cost per action
# --------------------------------------------------------------------------


def test_payload_cost_names_the_field_responsible_for_the_drain():
    """"This tool is expensive" is not actionable; naming the field is."""

    result = {
        "path": "x.docx",
        "table_count": 4,
        "tables": [{"cell_hashes": ["a" * 64] * 40}],
    }
    cost = payload_cost("audit_document", result)
    assert cost.dominant_field == "tables"
    assert cost.dominant_share > 0.9
    assert cost.estimated_tokens > 0


def test_payload_cost_handles_a_non_dict_result():
    cost = payload_cost("search_zotero", "plain text result")
    assert cost.dominant_field is None
    assert cost.dominant_share == 0.0


def test_turn_input_tokens_are_per_turn_not_cumulative():
    """The premise the payload fit rests on, pinned against the real log.

    Pricing a payload by subtracting turn 1 from turn 2 assumes the provider
    bills a running total. It does not: `list_bookmarks` reads fewer input
    tokens on turn 2 than on turn 1, which cumulative billing cannot produce.
    If this ever flips, `fit_payload_chars_per_token` is the wrong model.
    """

    turns = {
        record["turn"]: record["usage"]["input_tokens"]
        for record in _events("read_only")
        if record.get("event") == "model_response"
        and record.get("case") == "list_bookmarks"
        and record.get("toolset") == "minimal"
    }
    assert turns[2] < turns[1], turns


def test_payload_fit_separates_ratio_from_per_turn_overhead():
    small, large = (143, 372), (2333, 1984)
    ratio, overhead = fit_payload_chars_per_token([small, large])
    assert 1.2 < ratio < 1.5, ratio
    assert overhead > 0
    for chars, tokens in (small, large):
        assert overhead + chars / ratio == pytest.approx(tokens, rel=1e-6)


def test_payload_fit_holdout_residual_is_reported_honestly():
    """The fit is a two-point estimate, not a calibration; hold one out.

    A third recorded observation misses by roughly 9%. That is why payload
    cost is argued in characters and the token figure is quoted as a range.
    """

    ratio, overhead = fit_payload_chars_per_token([(143, 372), (2333, 1984)])
    holdout_chars, holdout_tokens = 150, 414
    predicted = overhead + holdout_chars / ratio
    residual = abs(predicted - holdout_tokens) / holdout_tokens
    assert 0.05 < residual < 0.15, residual


def test_payload_fit_refuses_input_it_cannot_separate():
    with pytest.raises(ValueError, match="at least two"):
        fit_payload_chars_per_token([(100, 50)])
    with pytest.raises(ValueError, match="differ in token count"):
        fit_payload_chars_per_token([(100, 50), (200, 50)])


# --------------------------------------------------------------------------
# Ease of use
# --------------------------------------------------------------------------


def test_interaction_quality_separates_recovery_from_first_call_success():
    events = [
        {"event": "tool_result", "case": "c", "toolset": "minimal", "repeat": 1,
         "name": "format_tables", "output": {"ok": False, "error": "bad spec"}},
        {"event": "tool_result", "case": "c", "toolset": "minimal", "repeat": 1,
         "name": "format_tables", "output": {"ok": False, "error": "bad spec"}},
        {"event": "tool_result", "case": "c", "toolset": "minimal", "repeat": 1,
         "name": "format_tables", "output": {"ok": True, "result": {}}},
    ]
    (item,) = interaction_quality(events)
    assert item.tool_calls == 3
    assert item.failed_calls == 2
    assert item.calls_before_first_success == 2
    assert item.recovered_after_failure is True
    assert item.first_call_success is False
    assert item.failing_tools == ["format_tables"]
