"""Plan 924: text typed at a zero-length anchor lands where Word itself types it (live; RESEARCH_MCP_WORD_LIVE=1).

The synthetic `POINT_BODY` (point comments, collapsed bookmarks, a point comment beside a collapsed bookmark) is
typed in a hidden, separate Word with Track Changes on (Selection.TypeText) and, on another copy, inserted with it
off (Range.InsertAfter); edit_document makes the same insertions tracked (research-mcp switch on) and plain
(through Word on this revision-free document). Each paragraph's marker order must equal Word's. Word renumbers
comment ids on save, so ids are compared by order of appearance.
"""
from __future__ import annotations

import os
import re
import shutil
from pathlib import Path

import pytest

from test_revisions import run_batch
from test_revisions_anchor_cases import POINT_BODY, POINT_CASES, POINT_COMMENTS, make, marked, point_ops

pytestmark = [
    pytest.mark.skipif(os.name != "nt", reason="needs Windows"),
    pytest.mark.skipif(os.environ.get("RESEARCH_MCP_WORD_LIVE") != "1", reason="needs RESEARCH_MCP_WORD_LIVE=1"),
]

_SKIP = {"\x05", "\x15", "\x13", "\x14", "\x01", "\r", "\x07"}


def _ids_by_order(text: str) -> str:
    order: dict[str, str] = {}
    return re.sub(r"(?<=[c^])(\d+)", lambda m: order.setdefault(m.group(1), str(len(order) + 1)), text)


def _word_insert(source: Path, out: Path, track: bool) -> None:
    import pythoncom
    import win32com.client

    shutil.copyfile(source, out)
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")  # its own process; never the user's Word
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(str(out), False, False, False)
        doc.TrackRevisions = track
        for index in sorted(POINT_CASES, reverse=True):
            text, at, _ = POINT_CASES[index]
            paragraph = doc.Paragraphs(index).Range
            starts = [int(item.Start) for item in paragraph.Characters if str(item.Text) not in _SKIP]
            spot = starts[at - 1] + 1 if at else int(paragraph.Start)
            target = doc.Range(spot, spot)
            if track:
                target.Select()
                word.Selection.TypeText(f"<{index}>")
            else:
                target.InsertAfter(f"<{index}>")
        doc.TrackRevisions = False
        doc.Save()
        doc.Close(False)
    finally:
        word.Quit(False)


@pytest.mark.parametrize("track", [True, False], ids=["tracked", "plain"])
def test_live_point_anchor_typing_matches_word(tmp_path, track):
    source = make(tmp_path, "points", POINT_BODY, POINT_COMMENTS)
    by_word = tmp_path / "word.docx"
    _word_insert(source, by_word, track)
    by_mcp = tmp_path / "mcp.docx"
    shutil.copyfile(source, by_mcp)
    pre = [{"op_id": "on", "type": "set_track_changes", "on": True, "author": "Reviewer B"}] if track else []
    result = run_batch(by_mcp, *pre, *point_ops())
    assert result.status == "complete", result
    for index, (_, _, expected) in POINT_CASES.items():
        wanted = expected.replace("{", "{+" if track else "").replace("}", "+}" if track else "")
        assert _ids_by_order(marked(by_word, index)) == _ids_by_order(wanted), f"Word, paragraph {index}"
        assert _ids_by_order(marked(by_mcp, index)) == _ids_by_order(wanted), f"MCP, paragraph {index}"
