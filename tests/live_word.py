"""Shared helper for live Word checks (RESEARCH_MCP_WORD_LIVE=1): open a document in a private, hidden Word
(`DispatchEx`, never the user's own Word), run Accept All and Reject All, and save each result. No document data
lives here, so public live tests can use it."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

COM_CHILD = r'''
import json, sys, pythoncom, win32com.client
jobs = json.loads(sys.stdin.read())  # on stdin: many long paths overflow the Windows command line
want_types = len(sys.argv) > 1  # also list each revision's WdRevisionType
pythoncom.CoInitialize()
app = win32com.client.DispatchEx("Word.Application")
results = []
def open_(path):
    return app.Documents.Open(FileName=path, ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False,
        Revert=False, Visible=False, OpenAndRepair=False, NoEncodingDialog=True)
try:
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        app.AutomationSecurity = 3
    except Exception:
        pass
    for job in jobs:
        src, acc, rej = job[:3]
        save = job[3] if len(job) > 3 else None
        out = {"opens_clean": False}
        try:
            doc = open_(src)
            out["word_revisions"] = doc.Revisions.Count
            out["word_comments"] = doc.Comments.Count
            if want_types:
                out["types"] = [rev.Type for rev in doc.Revisions]
            if save:  # Word's own save of the untouched output, opened again (a repair prompt raises here)
                doc.SaveAs2(FileName=save, FileFormat=16, AddToRecentFiles=False)
                doc.Close(SaveChanges=0)
                doc = open_(save)
                out["saved_revisions"] = doc.Revisions.Count
            doc.Close(SaveChanges=0)
            out["opens_clean"] = True
            for target, verb in ((acc, "AcceptAll"), (rej, "RejectAll")):
                doc = open_(src)
                getattr(doc.Revisions, verb)()
                out[verb + "_left"] = doc.Revisions.Count
                out[verb + "_comments"] = doc.Comments.Count
                doc.SaveAs2(FileName=target, FileFormat=16, AddToRecentFiles=False)
                doc.Close(SaveChanges=0)
        except Exception as exc:
            out["error"] = repr(exc)[:400]
        results.append(out)
finally:
    try:
        app.Quit(SaveChanges=0)
    except Exception:
        pass
print(json.dumps(results))
'''


def word_resolve(docs: list[Path], resave: bool = False, types: bool = False) -> list[dict]:
    """Open + Accept All + Reject All of each document in one hidden Word instance of our own. Each result has
    `opens_clean`, `word_revisions`, `word_comments`, `AcceptAll_left`/`RejectAll_left` (revisions left),
    `AcceptAll_comments`/`RejectAll_comments` and the saved paths `accept` and `reject`. `resave`: Word also saves
    each document untouched (`save`) and opens that save again (`saved_revisions`); `types`: each revision's
    WdRevisionType (`types`; 1 insert, 2 delete, 3 formatting)."""
    jobs = [(str(d), str(d.with_name(d.stem + "-word-accept.docx")), str(d.with_name(d.stem + "-word-reject.docx")))
            + ((str(d.with_name(d.stem + "-word-save.docx")),) if resave else ()) for d in docs]
    for job in jobs:
        for path in job[1:]:
            Path(path).unlink(missing_ok=True)
    r = subprocess.run([sys.executable, "-c", COM_CHILD] + (["types"] if types else []), input=json.dumps(jobs),
                       capture_output=True, text=True, timeout=900)
    line = (r.stdout.strip().splitlines() or [""])[-1]
    assert line.startswith("["), r.stderr[-800:]
    results = json.loads(line)
    for result, job in zip(results, jobs):
        result.update(accept=Path(job[1]), reject=Path(job[2]))
        if resave:
            result["save"] = Path(job[3])
    return results
