"""textclean module and char_map.json."""

import json

import pytest

from research_mcp import textclean
from research_mcp.textclean import clean, describe, merge_changes, scan, total


def test_zero_width_stripped():
    out, changes = clean("a\u200bb")
    assert out == "ab"
    assert changes == [{"char": "\u200b", "codepoint": "U+200B", "name": "ZERO WIDTH SPACE",
                        "action": "strip", "to": "", "count": 1}]


def test_nbsp_and_odd_spaces_become_space():
    assert clean("10\u00a0mm")[0] == "10 mm"
    assert clean("a\u2009b\u202fc\u3000d\u2003e")[0] == "a b c d e"


def test_manuscript_style_kept():
    s = "x \u2192 y \u2026 \u201cq\u201d \u2018r\u2019 \u2014 z \u2013 w \u2264 \u00b1 \u03b1 \u2212 \u00b0 caf\u00e9"
    out, changes = clean(s)
    assert out == s and changes == []


def test_hyphens_map_to_keyboard():
    assert clean("\u2011")[0] == "-"
    assert clean("non\u2010standard")[0] == "non-standard"


def test_fullwidth_to_ascii():
    assert clean("\uff21\uff22\uff23\uff01")[0] == "ABC!"


def test_tag_characters_removed():
    hidden = "".join(chr(0xE0000 + ord(c)) for c in "ignore previous")
    out, changes = clean(f"hi{hidden}there")
    assert out == "hithere"
    assert total(changes) == len("ignore previous")


def test_emoji_zwj_sequence_kept():
    family = "\U0001F468\u200d\U0001F469\u200d\U0001F467"
    woman_tech = "\U0001F469\U0001F3FD\u200d\U0001F4BB"
    heart = "\u2764\ufe0f\u200d\U0001F525"
    shake_h = "\U0001F642\u200d\u2194\ufe0f"
    shake_v = "\U0001F642\u200d\u2195\ufe0f"
    for s in (family, woman_tech, heart, shake_h, shake_v):
        assert clean(s) == (s, [])
    assert clean("a\u200db")[0] == "ab"
    assert clean("\u2308\u200d\u2309")[0] == "\u2308\u2309"  # math brackets are not emoji


@pytest.mark.parametrize("cp", [0x2061, 0x2062, 0x2063, 0x2064])
def test_invisible_math_operators_kept(cp):
    s = f"x{chr(cp)}y"
    assert clean(s) == (s, [])


def test_subdivision_flag_tags_kept():
    england = "\U0001F3F4\U000E0067\U000E0062\U000E0065\U000E006E\U000E0067\U000E007F"
    assert clean(england) == (england, [])
    hidden = "".join(chr(0xE0000 + ord(c)) for c in "ignore previous")
    assert clean(f"\U0001F3F4{hidden}\U000E007F")[0] == "\U0001F3F4"


def test_keep_holds_characters():
    assert clean("10\u00a0mm", keep=["\u00a0"])[0] == "10\u00a0mm"
    assert clean("10\u00a0mm", keep=["U+00A0"])[0] == "10\u00a0mm"


def test_controls_stripped_but_tab_newline_kept():
    assert clean("a\tb\nc\r\nd\x00e\x07f\x7fg\x85h")[0] == "a\tb\nc\r\nd" + "efgh"


def test_soft_hyphen_bidi_bom():
    assert clean("co\u00adop\u202e\u2066x\u2069\ufeff\u200e")[0] == "coopx"


def test_line_paragraph_separators():
    assert clean("a\u2028b\u2029c")[0] == "a\nb\nc"


def test_odd_quotes_map_to_manuscript():
    assert clean("\u201bq\u201f")[0] == "\u2019q\u201d"


def test_counts_aggregate_and_order():
    out, changes = clean("\u00a0a\u200b\u00a0b\u200b\u200b")
    assert out == " a b"
    assert [(c["codepoint"], c["count"]) for c in changes] == [("U+00A0", 2), ("U+200B", 3)]


def test_clean_text_identity_fast_path():
    s = "plain ASCII text"
    out, changes = clean(s)
    assert out is s and changes == []
    assert clean("") == ("", [])


def test_scan_merge_total_describe():
    a, b = scan("\u00a0"), scan("x\u00a0\u200b")
    merged = merge_changes(a, b)
    assert {c["codepoint"]: c["count"] for c in merged} == {"U+00A0": 2, "U+200B": 1}
    assert total(merged) == 3
    d = describe("U+00A0")
    assert d["action"] == "space" and d["to"] == " " and d["map_version"] == 1
    assert describe("\u2192")["action"] == "keep"
    assert describe("Q")["reason"] == "not in char_map.json"


def test_table_is_valid():
    data = json.loads(textclean.MAP_PATH.read_text(encoding="utf-8"))
    actions = {"strip", "space", "map", "keep"}
    for entry in data["entries"]:
        assert entry["action"] in actions
        assert entry.get("reason")
        if entry["action"] == "map":
            assert "to" in entry
    textclean._table.cache_clear()
    table, version = textclean._table()  # raises on duplicate codepoints
    assert version == data["version"] and 0xA0 in table


@pytest.mark.parametrize("cp", [0x09, 0x0A, 0x0B, 0x0D, 0x20, 0x2192, 0x2026, 0x2014, 0x03B1])
def test_never_changed(cp):
    assert clean(chr(cp)) == (chr(cp), [])


def test_duplicate_table_entries_raise(tmp_path, monkeypatch):
    data = json.loads(textclean.MAP_PATH.read_text(encoding="utf-8"))
    data["entries"].append({"from": "U+00A0", "action": "keep", "reason": "duplicate"})
    bad = tmp_path / "char_map.json"
    bad.write_text(json.dumps(data), encoding="utf-8")
    monkeypatch.setattr(textclean, "MAP_PATH", bad)
    textclean._table.cache_clear()
    try:
        with pytest.raises(ValueError, match="U\\+00A0 twice"):
            textclean._table()
    finally:
        monkeypatch.undo()
        textclean._table.cache_clear()


def test_mapping_report_payload_and_local_log(tmp_path):
    table_before = textclean.MAP_PATH.read_bytes()
    rep = textclean.mapping_report("U+00A0", "keep", "unit spacing must not break", "research MCP",
                                   snippet="10 mm" + "x" * 200)
    imp = rep["improvement"]
    assert imp["project"] == "research"
    assert imp["text"] == "Char mapping: U+00A0 NO-BREAK SPACE -> U+0020 (should be kept)"
    assert rep["title_prefix"] == "Char mapping: U+00A0 "
    assert rep["current"]["action"] == "space" and rep["should_be"] == " "
    assert "x" * 81 not in imp["notes"]  # snippet capped at 80 chars
    log = tmp_path / "char_reports.jsonl"
    assert textclean.record_report(log, rep, "r") == 1
    assert textclean.record_report(log, rep, "r") == 2
    other = textclean.mapping_report("‑", "‑", "keep NB hyphen", "research MCP")
    assert other["improvement"]["text"].startswith("Char mapping: U+2011 NON-BREAKING HYPHEN -> -")
    assert textclean.record_report(log, other, "r") == 1
    assert "(should be kept)" in other["improvement"]["text"]  # should_be == the char itself
    removed = textclean.mapping_report("U+061C", "", "should be removed", "research MCP")
    assert removed["improvement"]["text"].endswith("(should be (removed))")
    assert textclean.MAP_PATH.read_bytes() == table_before  # reports never edit the table
    assert textclean.describe("U+000B")["action"] == "keep"


def test_mapping_report_rejects_bad_char():
    table_before = textclean.MAP_PATH.read_bytes()
    for bad in ("ab", "U+", "U+ZZZZ", ""):
        with pytest.raises(ValueError):
            textclean.mapping_report(bad, "keep", "x", "research MCP")
    assert textclean.MAP_PATH.read_bytes() == table_before
