"""Clean invisible and unusual characters from text an agent writes (plan 551).

One pure-Python module and one table (``char_map.json`` next to this file).

    clean(text, keep=()) -> (cleaned_text, changes)

``changes`` is a list of ``{char, codepoint, name, action, to, count}`` in order of first appearance,
empty when nothing changed. Actions: ``strip`` (removed), ``space`` (became " "), ``map`` (became
``to``). ``keep`` is the per-call escape hatch: characters (or ``"U+00A0"`` strings) left alone.

Rules the table cannot express, handled here:
  - U+200D (zero-width joiner) is kept between two emoji, so emoji ZWJ sequences survive;
  - 3-7 tag letters/digits closed by U+E007F are kept right after U+1F3F4 (subdivision flags);
  - fullwidth ASCII (``to: "ascii"``) maps by offset to its ASCII counterpart.

Callers decide what is text: equations, field codes and Zotero fields are never passed in.
"""

from __future__ import annotations

import json
import unicodedata
from functools import lru_cache
from pathlib import Path
from typing import Iterable

MAP_PATH = Path(__file__).with_name("char_map.json")

_ZWJ = 0x200D
_BLACK_FLAG = 0x1F3F4
_CANCEL_TAG = 0xE007F
_VS16 = 0xFE0F


def _cp(value: str) -> int:
    return int(value.upper().removeprefix("U+"), 16)


@lru_cache(maxsize=1)
def _table() -> tuple[dict[int, dict], int]:
    data = json.loads(MAP_PATH.read_text(encoding="utf-8"))
    table: dict[int, dict] = {}
    for entry in data["entries"]:
        start = _cp(entry["from"])
        end = _cp(entry.get("through", entry["from"]))
        for cp in range(start, end + 1):
            if cp in table:
                raise ValueError(f"char_map.json lists U+{cp:04X} twice")
            table[cp] = entry
    return table, int(data.get("version", 1))


def map_version() -> int:
    return _table()[1]


def codepoint(char: str) -> str:
    return f"U+{ord(char):04X}"


def char_name(char: str) -> str:
    try:
        return unicodedata.name(char)
    except ValueError:
        return {0x9: "TAB", 0xA: "LINE FEED", 0xD: "CARRIAGE RETURN"}.get(ord(char), f"<control {codepoint(char)}>")


def entry_for(char: str) -> dict | None:
    """The table entry for one character, or None when it is not listed (kept)."""
    return _table()[0].get(ord(char))


def _keep_set(keep: Iterable[str]) -> set[int]:
    out: set[int] = set()
    for item in keep or ():
        if not item:
            continue
        if item.upper().startswith("U+") and len(item) > 2:
            out.add(_cp(item))
        else:
            out.update(ord(c) for c in item)
    return out


_EMOJI_BASES = frozenset(
    [0x00A9, 0x00AE, 0x203C, 0x2049, 0x2122, 0x2139, 0x2194, 0x2195, 0x2196, 0x2197, 0x2198, 0x2199,
     0x21A9, 0x21AA, 0x231A, 0x231B, 0x2328, 0x23CF, *range(0x23E9, 0x23F4), 0x23F8, 0x23F9, 0x23FA,
     0x24C2, 0x25AA, 0x25AB, 0x25B6, 0x25C0, *range(0x25FB, 0x25FF), 0x2B05, 0x2B06, 0x2B07, 0x2B1B,
     0x2B1C, 0x2B50, 0x2B55, 0x3030, 0x303D, 0x3297, 0x3299]
)


def _is_emoji(cp: int) -> bool:
    """Characters that can sit either side of a ZWJ in an emoji sequence (math symbols excluded)."""
    return 0x1F000 <= cp <= 0x1FAFF or 0x2600 <= cp <= 0x27BF or cp in _EMOJI_BASES


def _zwj_in_emoji(cps: list[int], i: int) -> bool:
    j = i - 1
    while j >= 0 and (cps[j] == _VS16 or 0x1F3FB <= cps[j] <= 0x1F3FF):
        j -= 1
    return j >= 0 and i + 1 < len(cps) and _is_emoji(cps[j]) and _is_emoji(cps[i + 1])


def _is_flag_tag(cp: int) -> bool:
    return 0xE0030 <= cp <= 0xE0039 or 0xE0061 <= cp <= 0xE007A  # tag digits and lower-case letters


def _flag_tags(cps: list[int]) -> set[int]:
    """Indexes of tag characters forming emoji subdivision flags (U+1F3F4, 3-7 tag letters/digits, U+E007F)."""
    keep: set[int] = set()
    for i, cp in enumerate(cps):
        if cp != _BLACK_FLAG:
            continue
        j = i + 1
        while j < len(cps) and _is_flag_tag(cps[j]):
            j += 1
        if 3 <= j - i - 1 <= 7 and j < len(cps) and cps[j] == _CANCEL_TAG:
            keep.update(range(i + 1, j + 1))
    return keep


def _replacement(cp: int, entry: dict) -> str | None:
    action = entry["action"]
    if action == "strip":
        return ""
    if action == "space":
        return " "
    if action == "map":
        to = entry["to"]
        if to == "ascii":
            return chr(cp - 0xFEE0)
        return to
    return None  # keep


def clean(text: str, keep: Iterable[str] = ()) -> tuple[str, list[dict]]:
    """Return ``(cleaned, changes)``. ``text`` that needs nothing comes back unchanged with ``[]``."""
    if not text:
        return text, []
    table = _table()[0]
    if text.isascii() and not any(ord(c) in table for c in text):
        return text, []
    kept = _keep_set(keep)
    cps = [ord(c) for c in text]
    flag_tags = _flag_tags(cps) if any(0xE0000 <= cp <= 0xE007F for cp in cps) else set()
    out: list[str] = []
    changes: dict[int, dict] = {}
    for i, cp in enumerate(cps):
        entry = table.get(cp)
        if entry is None or cp in kept or i in flag_tags or (cp == _ZWJ and _zwj_in_emoji(cps, i)):
            out.append(chr(cp))
            continue
        repl = _replacement(cp, entry)
        if repl is None:
            out.append(chr(cp))
            continue
        out.append(repl)
        rec = changes.get(cp)
        if rec is None:
            ch = chr(cp)
            changes[cp] = {
                "char": ch,
                "codepoint": codepoint(ch),
                "name": char_name(ch),
                "action": entry["action"],
                "to": repl,
                "count": 1,
            }
        else:
            rec["count"] += 1
    if not changes:
        return text, []
    return "".join(out), list(changes.values())


def scan(text: str, keep: Iterable[str] = ()) -> list[dict]:
    """The changes ``clean`` would make, without the text (for inspect counts)."""
    return clean(text, keep)[1]


def merge_changes(*lists: Iterable[dict]) -> list[dict]:
    """Merge change lists from several strings into one, summing counts per codepoint."""
    merged: dict[str, dict] = {}
    for changes in lists:
        for rec in changes or ():
            cur = merged.get(rec["codepoint"])
            if cur is None:
                merged[rec["codepoint"]] = dict(rec)
            else:
                cur["count"] += rec["count"]
    return list(merged.values())


def total(changes: Iterable[dict]) -> int:
    return sum(rec["count"] for rec in changes or ())


def describe(char_or_cp: str) -> dict:
    """What the table does with one character: for wrong-mapping reports."""
    ch = chr(_cp(char_or_cp)) if char_or_cp.upper().startswith("U+") and len(char_or_cp) > 2 else char_or_cp[:1]
    entry = entry_for(ch)
    if entry is None:
        return {"char": ch, "codepoint": codepoint(ch), "name": char_name(ch), "action": "keep", "to": ch,
                "reason": "not in char_map.json", "map_version": map_version()}
    repl = _replacement(ord(ch), entry)
    return {"char": ch, "codepoint": codepoint(ch), "name": char_name(ch), "action": entry["action"],
            "to": ch if repl is None else repl, "reason": entry.get("reason", ""), "map_version": map_version()}


# --- Wrong-mapping reports (plan 551 S4) ------------------------------------------------------------
# MCP processes don't post anywhere: a report is logged locally and comes back as a ready payload. By
# default that is a prefilled GitHub issue link (server.report_issue). With RESEARCH_MCP_ISSUES=agent it is
# an improvement payload plus ``file_with`` steps the agent follows with its own issue tool, adding a note
# to an open item with the same title prefix instead of filing a duplicate.
# Reports never edit char_map.json; the table changes only after review.

SNIPPET_MAX = 80


def _one_char(value: str) -> str:
    value = (value or "").strip() if len((value or "").strip()) > 1 else (value or "")
    if value.upper().startswith("U+") and len(value) > 2:
        return chr(_cp(value))
    if len(value) != 1:
        raise ValueError("char must be one character or a U+XXXX code point")
    return value


def _show(text: str) -> str:
    if text == "":
        return "(removed)"
    return " ".join(f"U+{ord(c):04X}" if c.isspace() or not c.isprintable() else c for c in text)


def mapping_report(char: str, should_be: str, reason: str, mcp: str, snippet: str | None = None,
                   mapped_to: str | None = None) -> dict:
    """Build a wrong-mapping report: the table's current treatment plus an improvement payload (text, notes).

    ``should_be`` is what the character should have become ("keep" to leave it alone). ``snippet`` is
    at most SNIPPET_MAX characters the model chooses (never a whole paragraph or slide).
    """
    ch = _one_char(char)
    current = describe(ch)
    seen = current["to"] if mapped_to is None else mapped_to
    want = ch if should_be.strip().lower() == "keep" else should_be
    snip = (snippet or "")[:SNIPPET_MAX]
    title = (f"Char mapping: {current['codepoint']} {current['name']} -> {_show(seen)} "
             f"(should be {'kept' if want == ch else _show(want)})")
    notes = "\n".join([
        f"Reported through the {mcp} (report_issue category char_mapping).",
        f"Table: char_map.json version {current['map_version']}, action {current['action']}, reason: {current['reason']}.",
        f"Seen: {_show(seen)}; should be: {'kept as is' if want == ch else _show(want)}.",
        f"Reason given: {reason.strip()[:1000]}",
        *( [f"Snippet ({len(snip)} chars, chosen by the model): {snip!r}"] if snip else [] ),
        "Fix: edit char_map.json and extend tests/test_textclean.py.",
    ])
    return {
        "codepoint": current["codepoint"],
        "name": current["name"],
        "current": {"action": current["action"], "to": current["to"], "reason": current["reason"]},
        "should_be": want,
        "improvement": {"text": title[:500], "notes": notes[:4000], "project": "research"},
        "title_prefix": f"Char mapping: {current['codepoint']} ",
    }


def record_report(path, report: dict, reason: str) -> int:
    """Append a report to a local JSONL; return how many reports this code point now has there."""
    from datetime import datetime, timezone

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    count = 1
    if path.exists():
        for line in path.read_text(encoding="utf-8").splitlines():
            try:
                count += json.loads(line).get("codepoint") == report["codepoint"]
            except ValueError:
                continue
    entry = {"at": datetime.now(timezone.utc).isoformat(timespec="seconds"), "codepoint": report["codepoint"],
             "name": report["name"], "current": report["current"], "should_be": report["should_be"],
             "reason": reason.strip()[:1000], "count": count}
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=True) + "\n")
    return count


FILE_WITH = ("File it with the issue or task tool you have. 1) Look for an open item whose text starts with "
             "title_prefix. 2) If one exists, add improvement.notes to it as a note. 3) Otherwise create an "
             "item from improvement.text and improvement.notes. Without such a tool, tell the user; the "
             "report is already logged locally.")
