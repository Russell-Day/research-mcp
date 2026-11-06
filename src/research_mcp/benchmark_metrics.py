"""Read-only cost analysis over benchmark artifacts.

The agent benchmark reports one token number per case. That number answers
"what did this run cost" but not the two questions that decide whether an MCP
surface is usable by a model:

1. **Context inclusion** -- what does merely *attaching* the server cost, on
   every turn, before the agent does anything?
2. **Cost per action** -- what does one tool call add to the context, and does
   it stay there?

Both are measurable from artifacts already on disk. Nothing here runs a model,
opens Word, touches Zotero, or takes the integration lease, so it is safe to
run while another agent is benchmarking.

Everything is a pure function over ``benchmark_runs/<stamp>/events.jsonl`` or
over a tool-result object, so the runner itself is not modified.
"""

from __future__ import annotations

import argparse
from collections import defaultdict
from dataclasses import dataclass, field
import json
from pathlib import Path
from typing import Any, Iterable


# Schema text and hash-dense result payloads tokenize very differently, so one
# shared ratio would misreport one of them by more than a factor of two.
#
# The two numbers do NOT have equal standing, and callers should not treat them
# as if they do:
#
# * SCHEMA_CHARS_PER_TOKEN is calibrated across five independent cases whose
#   derived ratios span 0.04 (see `calibrate_chars_per_token`). It is a
#   measurement.
# * PAYLOAD_CHARS_PER_TOKEN is a two-point fit with one holdout that missed by
#   ~9% (see `fit_payload_chars_per_token`). It is an estimate, and payload
#   sizes should be reported in characters wherever the argument has to hold.
SCHEMA_CHARS_PER_TOKEN = 3.72
PAYLOAD_CHARS_PER_TOKEN = 1.36


def load_events(run_dir: Path) -> list[dict[str, Any]]:
    path = Path(run_dir) / "events.jsonl"
    events: list[dict[str, Any]] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            record = json.loads(line)
        except ValueError:
            continue  # a truncated final line must not discard the whole run
        if isinstance(record, dict):
            events.append(record)
    return events


def json_chars(value: Any) -> int:
    return len(json.dumps(value, ensure_ascii=False, default=str))


# --------------------------------------------------------------------------
# Axis 1: context inclusion
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class SchemaCost:
    """Exact token cost of the tool-schema block, in the model's own tokenizer.

    Turn 1 carries no conversation history in either condition, and both
    conditions replay an identical system prompt and user prompt, so the
    difference between the two turn-1 `input_tokens` readings isolates the
    schemas that `minimal` withheld. No tokenizer library is involved and
    nothing is estimated: the provider counted these.
    """

    case_id: str
    repeat: int
    all_first_turn_tokens: int
    minimal_first_turn_tokens: int

    @property
    def withheld_schema_tokens(self) -> int:
        return self.all_first_turn_tokens - self.minimal_first_turn_tokens

    @property
    def schema_share_of_first_turn(self) -> float:
        if self.all_first_turn_tokens <= 0:
            return 0.0
        return self.withheld_schema_tokens / self.all_first_turn_tokens


def _first_turn_tokens(events: Iterable[dict[str, Any]]) -> dict[tuple[str, str, int], int]:
    seen: dict[tuple[str, str, int], int] = {}
    for record in events:
        if record.get("event") != "model_response" or record.get("turn") != 1:
            continue
        key = (str(record.get("case")), str(record.get("toolset")), int(record.get("repeat") or 1))
        usage = record.get("usage") or {}
        if key not in seen:
            seen[key] = int(usage.get("input_tokens") or 0)
    return seen


def schema_context_cost(events: Iterable[dict[str, Any]]) -> list[SchemaCost]:
    """Pair each case's `all` and `minimal` turn-1 readings."""

    first = _first_turn_tokens(events)
    paired: list[SchemaCost] = []
    for (case_id, toolset, repeat), tokens in sorted(first.items()):
        if toolset != "all":
            continue
        minimal = first.get((case_id, "minimal", repeat))
        if minimal is None:
            continue
        paired.append(
            SchemaCost(
                case_id=case_id,
                repeat=repeat,
                all_first_turn_tokens=tokens,
                minimal_first_turn_tokens=minimal,
            )
        )
    return paired


def calibrate_chars_per_token(
    costs: Iterable[SchemaCost], schema_chars: dict[str, int]
) -> dict[str, float]:
    """Derive chars-per-token from measured deltas and serialized schema size.

    `schema_chars` maps a tool name to the serialized length of the schema the
    runner sends. A case whose `minimal` condition exposed only its own
    required tool withheld ``total - required`` characters, and the provider
    reported exactly how many tokens that was. The resulting ratio predicts the
    per-turn context cost of a tool that does not exist yet, offline, from its
    schema alone.
    """

    total = sum(schema_chars.values())
    ratios: dict[str, float] = {}
    for cost in costs:
        withheld_chars = total - schema_chars.get(cost.case_id, 0)
        if cost.withheld_schema_tokens > 0:
            ratios[cost.case_id] = withheld_chars / cost.withheld_schema_tokens
    return ratios


def predict_schema_tokens(schema: Any, *, chars_per_token: float = SCHEMA_CHARS_PER_TOKEN) -> int:
    """Predicted per-turn context cost of adding one tool, before it is built."""

    return round(json_chars(schema) / chars_per_token)


def fit_payload_chars_per_token(
    observations: Iterable[tuple[int, int]]
) -> tuple[float, float]:
    """Fit ``tokens = overhead + chars / ratio`` to (payload chars, turn tokens).

    Result payloads cannot be priced by subtraction the way schemas can. This
    provider reports each turn's *new* input rather than a running total --
    `list_bookmarks` reads 459 tokens on turn 1 and 372 on turn 2, which is
    impossible under cumulative billing -- so a turn-2 reading is one payload
    plus a fixed per-turn overhead, not a delta against turn 1.

    Two observations far apart in size separate the two unknowns. Callers
    should hold a third out and check the residual before trusting the ratio;
    on the recorded runs the holdout missed by about 9%, which is why payload
    cost is reported in characters and the token figure is given as a range.
    """

    points = sorted(observations)
    if len(points) < 2:
        raise ValueError("need at least two observations to separate ratio from overhead")
    (small_chars, small_tokens), (large_chars, large_tokens) = points[0], points[-1]
    if large_tokens == small_tokens:
        raise ValueError("observations must differ in token count")
    ratio = (large_chars - small_chars) / (large_tokens - small_tokens)
    overhead = small_tokens - small_chars / ratio
    return ratio, overhead


# --------------------------------------------------------------------------
# Axis 2: cost per action
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class PayloadCost:
    tool: str
    chars: int
    estimated_tokens: int
    dominant_field: str | None
    dominant_share: float


def payload_cost(
    tool: str, result: Any, *, chars_per_token: float = PAYLOAD_CHARS_PER_TOKEN
) -> PayloadCost:
    """Cost of one tool result, and which field is responsible for it.

    A tool result is not a transient cost: it is appended to the conversation
    and every later turn carries it. Naming the dominant field turns "this tool
    is expensive" into an actionable schema change.
    """

    chars = json_chars(result)
    dominant: str | None = None
    dominant_chars = 0
    if isinstance(result, dict):
        for key, value in result.items():
            size = json_chars(value)
            if size > dominant_chars:
                dominant, dominant_chars = str(key), size
    return PayloadCost(
        tool=tool,
        chars=chars,
        estimated_tokens=round(chars / chars_per_token),
        dominant_field=dominant,
        dominant_share=(dominant_chars / chars) if chars else 0.0,
    )


def payload_costs_from_events(events: Iterable[dict[str, Any]]) -> list[PayloadCost]:
    costs: list[PayloadCost] = []
    for record in events:
        if record.get("event") != "tool_result":
            continue
        payload = record.get("output") or {}
        if not payload.get("ok"):
            continue
        costs.append(payload_cost(str(record.get("name")), payload.get("result")))
    return costs


# --------------------------------------------------------------------------
# Ease of use
# --------------------------------------------------------------------------


@dataclass
class InteractionQuality:
    """Signals separating "the model succeeded" from "the tool was usable".

    A case that passes after four rejected argument shapes and one wrong tool
    is not the same product as one that passes on the first call, but the
    benchmark's status field records both as `passed`.
    """

    case_id: str
    toolset: str
    repeat: int
    tool_calls: int = 0
    failed_calls: int = 0
    distinct_tools_called: int = 0
    unrequired_tools_called: int = 0
    calls_before_first_success: int = 0
    recovered_after_failure: bool = False
    calls_after_first_success: int = 0
    failing_tools: list[str] = field(default_factory=list)

    @property
    def first_call_success(self) -> bool:
        return self.tool_calls > 0 and self.calls_before_first_success == 0

    def to_dict(self) -> dict[str, Any]:
        return {**self.__dict__, "first_call_success": self.first_call_success}


def interaction_quality(
    events: Iterable[dict[str, Any]],
    required_tools: dict[str, tuple[str, ...]] | None = None,
) -> list[InteractionQuality]:
    """Per case/toolset call-level usability, from `tool_result` events alone."""

    required_tools = required_tools or {}
    grouped: dict[tuple[str, str, int], list[dict[str, Any]]] = defaultdict(list)
    for record in events:
        if record.get("event") != "tool_result":
            continue
        key = (
            str(record.get("case")),
            str(record.get("toolset")),
            int(record.get("repeat") or 1),
        )
        grouped[key].append(record)

    report: list[InteractionQuality] = []
    for (case_id, toolset, repeat), records in sorted(grouped.items()):
        item = InteractionQuality(case_id=case_id, toolset=toolset, repeat=repeat)
        required = set(required_tools.get(case_id, ()))
        names: list[str] = []
        first_success: int | None = None
        for index, record in enumerate(records):
            name = str(record.get("name"))
            payload = record.get("output") or {}
            names.append(name)
            item.tool_calls += 1
            if payload.get("ok"):
                if first_success is None:
                    first_success = index
            else:
                item.failed_calls += 1
                if name not in item.failing_tools:
                    item.failing_tools.append(name)
        item.distinct_tools_called = len(set(names))
        if required:
            item.unrequired_tools_called = len(set(names) - required)
        if first_success is None:
            item.calls_before_first_success = item.tool_calls
        else:
            item.calls_before_first_success = first_success
            item.recovered_after_failure = first_success > 0
            item.calls_after_first_success = item.tool_calls - first_success - 1
        report.append(item)
    return report


# --------------------------------------------------------------------------
# Reporting
# --------------------------------------------------------------------------


def render_report(run_dirs: Iterable[Path]) -> str:
    run_dirs = list(run_dirs)
    lines = [
        "# MCP cost profile",
        "",
        "Generated from: " + ", ".join(f"`{Path(item).name}`" for item in run_dirs) + ".",
        "",
        "These are recorded runs. Payload sizes below describe the tool results as",
        "they were at the time of each run, so a later change to what a tool returns",
        "will not appear here until the run is repeated.",
        "",
        "## Axis 1: context inclusion (exact, provider-counted)",
        "",
        "Turn 1 carries no history, so `all` minus `minimal` is the schema block.",
        "",
        "| Run | Case | all t1 | minimal t1 | schema tokens | schema share of t1 |",
        "|---|---|---:|---:|---:|---:|",
    ]
    for run_dir in run_dirs:
        for cost in schema_context_cost(load_events(run_dir)):
            lines.append(
                f"| `{Path(run_dir).name}` | `{cost.case_id}` | {cost.all_first_turn_tokens} | "
                f"{cost.minimal_first_turn_tokens} | {cost.withheld_schema_tokens} | "
                f"{cost.schema_share_of_first_turn:.0%} |"
            )

    lines += [
        "",
        "## Axis 2: cost per action (tool-result payloads)",
        "",
        "A tool result is permanent context, not a transient cost.",
        "",
        "| Tool | Result chars | Est. tokens | Dominant field | Share |",
        "|---|---:|---:|---|---:|",
    ]
    seen: set[tuple[str, int]] = set()
    for run_dir in run_dirs:
        for cost in payload_costs_from_events(load_events(run_dir)):
            key = (cost.tool, cost.chars)
            if key in seen:
                continue
            seen.add(key)
            lines.append(
                f"| `{cost.tool}` | {cost.chars:,} | {cost.estimated_tokens:,} | "
                f"`{cost.dominant_field}` | {cost.dominant_share:.0%} |"
            )

    lines += [
        "",
        "## Ease of use",
        "",
        "| Run | Case | Toolset | Calls | Failed | Calls before 1st success | Recovered | Calls after 1st success |",
        "|---|---|---|---:|---:|---:|---|---:|",
    ]
    for run_dir in run_dirs:
        for item in interaction_quality(load_events(run_dir)):
            lines.append(
                f"| `{Path(run_dir).name}` | `{item.case_id}` | {item.toolset} | "
                f"{item.tool_calls} | {item.failed_calls} | {item.calls_before_first_success} | "
                f"{'yes' if item.recovered_after_failure else 'no'} | "
                f"{item.calls_after_first_success} |"
            )
    return "\n".join(lines) + "\n"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Read-only cost profile over benchmark runs")
    parser.add_argument("run_dirs", nargs="+", help="benchmark_runs/<timestamp> directories")
    parser.add_argument("--output", help="write the report here instead of stdout")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    dirs = [Path(item) for item in args.run_dirs if (Path(item) / "events.jsonl").exists()]
    if not dirs:
        raise SystemExit("no run directory contained events.jsonl")
    report = render_report(dirs)
    if args.output:
        Path(args.output).write_text(report, encoding="utf-8")
        print(args.output)
    else:
        print(report)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
