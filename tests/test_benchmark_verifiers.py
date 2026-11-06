"""Negative tests for the benchmark's scoring layer.

Every benchmark number depends on `verify_case` refusing a document that does
not satisfy the task. The existing suite exercises the agent loop thoroughly
but only ever drives the `tool_result` verifier kind through a fake session, so
`bookmark_cleanup`, `table_format`, and `zotero_citation` -- the three kinds
that inspect real DOCX packages -- have no test that they ever say *no*.

A verifier that cannot fail turns the whole benchmark into a transcript of the
model's own claims, which is exactly what the design set out to avoid.

Two tests here are `xfail(strict=True)`: they describe behaviour the verifier
should have, fail against the current implementation, and will turn the suite
red the moment the defect is fixed so the marker gets removed. Both defects are
documented in `planning/AGENT_BENCHMARK_SUITE.md` and
`planning/DECISIONS_AND_RATIONALE.md`.

All of this is offline: no model, no Word, no Zotero, no integration lease.
"""

from __future__ import annotations

import asyncio
import json
from pathlib import Path
import shutil
import zipfile

import pytest

from research_mcp.audit import audit_docx
from research_mcp.benchmark import BenchmarkCase, load_cases, verify_case
from research_mcp.server import server


# Packages are assembled directly rather than through `create_table_fixture`,
# which drives Word over COM. The verifiers under test only ever read OOXML, so
# building the bytes keeps this file genuinely offline -- which matters while a
# second agent is running live Word and Zotero benchmarks.
CONTENT_TYPES = (
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
)
DOCUMENT = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    "<w:body>%s<w:sectPr/></w:body></w:document>"
)
TABLE = (
    "<w:tbl><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>"
    "<w:tr><w:tc><w:p><w:r><w:t>LoRA</w:t></w:r></w:p></w:tc>"
    "<w:tc><w:p><w:r><w:t>Early</w:t></w:r></w:p></w:tc></w:tr>"
    "<w:tr><w:tc><w:p><w:r><w:t>LoRA</w:t></w:r></w:p></w:tc>"
    "<w:tc><w:p><w:r><w:t>Middle</w:t></w:r></w:p></w:tc></w:tr>"
    "</w:tbl>"
)


def _paragraph(text: str, bookmarks: tuple[str, ...] = ()) -> str:
    starts = "".join(
        f'<w:bookmarkStart w:id="{index}" w:name="{name}"/>'
        for index, name in enumerate(bookmarks, start=1)
    )
    ends = "".join(f'<w:bookmarkEnd w:id="{index}"/>' for index, _ in enumerate(bookmarks, start=1))
    return f"<w:p>{starts}<w:r><w:t>{text}</w:t></w:r>{ends}</w:p>"


def _package(path: Path, body: str) -> Path:
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("word/document.xml", DOCUMENT % body)
    return path


def _bookmark_package(path: Path, bookmarks: tuple[str, ...]) -> Path:
    """Identical body text in every variant, so only the bookmarks differ."""

    return _package(path, _paragraph("Synthetic table fixture.", bookmarks) + TABLE)


CITED_KEY = "6W98BDK7"
UNCITED_KEY = "ZZZZ9999"
# Copied out of the scratch `benchmark_runs/` namespace on purpose: these
# tests reproduce two live defects, and a skipped reproduction tracks nothing.
CITATION_ARTIFACT = Path(__file__).resolve().parent / "data" / "zotero_citation_verified.docx"

requires_citation_artifact = pytest.mark.skipif(
    not CITATION_ARTIFACT.exists(),
    reason="requires tests/data/zotero_citation_verified.docx",
)


def test_citation_artifact_fixture_is_present():
    """Unskipped: losing the fixture must fail, not silently skip the defect."""

    assert CITATION_ARTIFACT.exists(), CITATION_ARTIFACT


def _transcript(tool: str) -> list[dict]:
    """A minimal successful call, so verification depends only on the file."""

    return [{"kind": "tool", "name": tool, "output": {"ok": True, "result": {}}}]


def _case(tool: str, tier: str) -> BenchmarkCase:
    return BenchmarkCase(
        id="probe", description="", tier=tier, required_tools=(tool,), prompt="", verifier={}
    )


def _with_extra_body_text(source: Path, destination: Path, text: str) -> Path:
    """Copy a DOCX, adding one plain paragraph. No field is created.

    This is how a stray item key realistically reaches `document.xml`: quoted
    in prose, left behind by a donor, or echoed from the prompt.
    """

    with zipfile.ZipFile(source) as archive:
        names = archive.namelist()
        payload = {name: archive.read(name) for name in names}
    document = payload["word/document.xml"].decode("utf-8")
    assert text not in document, "probe text must not already be present"
    payload["word/document.xml"] = document.replace(
        "</w:body>", f"<w:p><w:r><w:t>see {text}</w:t></w:r></w:p></w:body>", 1
    ).encode("utf-8")
    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as out:
        for name in names:
            out.writestr(name, payload[name])
    return destination


# --------------------------------------------------------------------------
# zotero_citation
# --------------------------------------------------------------------------


@requires_citation_artifact
def test_zotero_citation_verifier_accepts_the_expected_item():
    """Positive control: the probe rig itself must be sound."""

    verified, info = verify_case(
        _case("insert_zotero_citation", "zotero_live"),
        {
            "kind": "zotero_citation",
            "tool": "insert_zotero_citation",
            "output": str(CITATION_ARTIFACT),
            "item_key": CITED_KEY,
        },
        _transcript("insert_zotero_citation"),
    )
    assert verified, info


@requires_citation_artifact
def test_zotero_citation_verifier_rejects_a_citation_of_the_wrong_item(tmp_path):
    """The most consequential false positive available.

    The document holds exactly one Zotero field, and it cites `6W98BDK7`. The
    requested key appears only as body text. All three checks still pass:
    `one_item_field` counts fields without reading them, `expected_item_key`
    greps the whole package, and `rendered` proves some field is not a
    placeholder without proving *which* item it renders.

    So a batch insert that cited the wrong paper would score as a pass -- and
    `insert_zotero_citations` is next on the build order.
    """

    probe = _with_extra_body_text(CITATION_ARTIFACT, tmp_path / "wrong_item.docx", UNCITED_KEY)
    verified, info = verify_case(
        _case("insert_zotero_citation", "zotero_live"),
        {
            "kind": "zotero_citation",
            "tool": "insert_zotero_citation",
            "output": str(probe),
            "item_key": UNCITED_KEY,
        },
        _transcript("insert_zotero_citation"),
    )
    assert not verified, info


@requires_citation_artifact
def test_zotero_citation_verifier_requires_the_output_to_exist(tmp_path):
    verified, info = verify_case(
        _case("insert_zotero_citation", "zotero_live"),
        {
            "kind": "zotero_citation",
            "tool": "insert_zotero_citation",
            "output": str(tmp_path / "never_written.docx"),
            "item_key": CITED_KEY,
        },
        _transcript("insert_zotero_citation"),
    )
    assert not verified
    assert "does not exist" in info["reason"]


# --------------------------------------------------------------------------
# bookmark_cleanup
# --------------------------------------------------------------------------


def _bookmark_verifier(source: Path, output: Path) -> dict:
    return {
        "kind": "bookmark_cleanup",
        "tool": "clean_bookmarks",
        "source": str(source),
        "output": str(output),
        "absent": ["MCPFixtureLegacy"],
        "present": ["_FixtureLegitimate"],
    }


def test_bookmark_cleanup_verifier_accepts_a_correct_cleanup(tmp_path):
    source = _bookmark_package(
        tmp_path / "source.docx", ("MCPFixtureLegacy", "_FixtureLegitimate")
    )
    output = _bookmark_package(tmp_path / "output.docx", ("_FixtureLegitimate",))

    verified, info = verify_case(
        _case("clean_bookmarks", "versioned"),
        _bookmark_verifier(source, output),
        _transcript("clean_bookmarks"),
    )
    assert verified, info


def test_bookmark_cleanup_verifier_rejects_a_cleanup_that_removed_too_much(tmp_path):
    """The check that does real work: a protected bookmark must survive."""

    source = _bookmark_package(
        tmp_path / "source.docx", ("MCPFixtureLegacy", "_FixtureLegitimate")
    )
    output = _bookmark_package(tmp_path / "output.docx", ())

    verified, info = verify_case(
        _case("clean_bookmarks", "versioned"),
        _bookmark_verifier(source, output),
        _transcript("clean_bookmarks"),
    )
    assert not verified
    assert info["checks"]["present"] is False


def test_bookmark_cleanup_verifier_detects_a_source_the_agent_overwrote(tmp_path):
    """Source-overwrite protection is the suite's central safety claim.

    Core and the benchmark write guard both refuse an overwrite, so this check
    is the third layer -- but as written it is decorative, and a regression in
    either of the other two would score as a clean pass.
    """

    source = _bookmark_package(
        tmp_path / "source.docx", ("MCPFixtureLegacy", "_FixtureLegitimate")
    )
    output = _bookmark_package(tmp_path / "output.docx", ("_FixtureLegitimate",))

    # The replacement keeps the body text byte-identical and only changes the
    # bookmark set. That isolates the defect: `body_text_unchanged`, `absent`
    # and `present` all still hold, so `source_hash_unchanged` is the only
    # check that could notice, and it cannot.
    original = audit_docx(source)
    clobbered = _bookmark_package(tmp_path / "clobbered.docx", ("MCPFixtureLegacy",))
    shutil.copy(clobbered, source)  # the agent overwrote its own source
    replaced = audit_docx(source)
    assert replaced.sha256 != original.sha256
    assert replaced.body_text_sha256 == original.body_text_sha256

    verifier = _bookmark_verifier(source, output)
    verifier["_source_sha256_before"] = original.sha256
    verified, info = verify_case(
        _case("clean_bookmarks", "versioned"),
        verifier,
        _transcript("clean_bookmarks"),
    )
    assert not verified, info


# --------------------------------------------------------------------------
# table_format
# --------------------------------------------------------------------------


def test_table_format_verifier_rejects_a_document_with_no_merges(tmp_path):
    """An untouched copy must not satisfy a spec that demands merges."""

    source = _bookmark_package(tmp_path / "source.docx", ())
    output = tmp_path / "output.docx"
    shutil.copy(source, output)

    verified, info = verify_case(
        _case("format_tables", "versioned"),
        {
            "kind": "table_format",
            "tool": "format_tables",
            "source": str(source),
            "output": str(output),
            "table_specs": [
                {
                    "index": 1,
                    "descriptive_columns": [1, 2],
                    "merges": [
                        {"column": 1, "start_row": 2, "end_row": 4, "label": "LoRA"}
                    ],
                }
            ],
        },
        _transcript("format_tables"),
    )
    assert not verified
    assert info["checks"]["merges_exact"] is False


# --------------------------------------------------------------------------
# Case catalog
# --------------------------------------------------------------------------


def test_verifier_refuses_an_unknown_kind(tmp_path):
    output = _bookmark_package(tmp_path / "out.docx", ())
    verified, info = verify_case(
        _case("audit_document", "read_only"),
        {"kind": "not_a_real_kind", "tool": "audit_document", "output": str(output)},
        _transcript("audit_document"),
    )
    assert not verified
    assert "unknown verifier kind" in info["reason"]


def test_verifier_refuses_when_the_required_tool_never_succeeded():
    verified, info = verify_case(
        _case("audit_document", "read_only"),
        {"kind": "tool_result", "tool": "audit_document", "contains": {"table_count": 1}},
        [{"kind": "tool", "name": "audit_document", "output": {"ok": False, "error": "boom"}}],
    )
    assert not verified
    assert "no successful" in info["reason"]


def test_every_case_is_either_exposed_by_the_server_or_marked_planned():
    """Self-maintaining replacement for a hardcoded list of planned case ids.

    As `inspect_document` and `edit_document` land, their cases must move out
    of the `planned` tier on their own; a literal id list silently rots.
    """

    exposed = {tool.name for tool in asyncio.run(server.list_tools())}
    cases = load_cases(Path("benchmarks/cases.json"))
    assert cases

    for case in cases:
        missing = set(case.required_tools) - exposed
        if case.tier == "planned":
            assert missing, (
                f"case {case.id!r} is still tier 'planned' but every required tool "
                f"is now exposed; promote it to an executable tier"
            )
        else:
            assert not missing, (
                f"case {case.id!r} is tier {case.tier!r} but the server does not "
                f"expose {sorted(missing)}; it would report 'unsupported' forever"
            )


def test_executable_cases_declare_a_real_verifier():
    """A `planned` verifier on an executable tier would auto-pass nothing."""

    for case in load_cases(Path("benchmarks/cases.json")):
        kind = case.verifier.get("kind")
        if case.tier == "planned":
            assert kind == "planned", case.id
        else:
            assert kind in {"tool_result", "bookmark_cleanup", "table_format", "zotero_citation", "document_edit"}, (
                f"case {case.id!r} has unrecognised verifier kind {kind!r}"
            )


def test_case_prompts_only_reference_variables_the_runner_supplies():
    """A typo'd placeholder fails the whole case at render time, mid-run."""

    supported = {
        "workspace",
        "run_dir",
        "fixture",
        "output",
        "citation_source",
        "comment_source",
        "item_donor",
        "bibliography_donor",
        "variant",
    }
    import string

    def placeholders(value: object) -> set[str]:
        # Mirror the runner's `_render`, which formats each string leaf rather
        # than the serialized structure; JSON braces are not placeholders.
        if isinstance(value, str):
            return {field for _, field, _, _ in string.Formatter().parse(value) if field}
        if isinstance(value, dict):
            return set().union(*(placeholders(item) for item in value.values())) if value else set()
        if isinstance(value, list):
            return set().union(*(placeholders(item) for item in value)) if value else set()
        return set()

    for case in load_cases(Path("benchmarks/cases.json")):
        names = placeholders(case.prompt) | placeholders(case.verifier)
        assert names <= supported, f"case {case.id!r} uses unknown placeholders {names - supported}"


@pytest.mark.skipif(
    not (Path("benchmark_runs/g12-versioned/format_tables--all-r1.docx").exists()),
    reason="requires a recorded merged format_tables output",
)
def test_table_format_verifier_accepts_a_correct_vertical_merge():
    """A correct merge must not score as a failure.

    `body_text` includes table cell paragraphs, and a vertical merge empties
    the cells it merges away, so the old `body_text_unchanged` check marked
    every correct merge as changed body text. The logical content is what
    matters and `cell_text_unchanged` already proves it.
    """

    source = Path("benchmark_runs/g12-versioned/word_fixture.docx")
    output = Path("benchmark_runs/g12-versioned/format_tables--all-r1.docx")
    before, after = audit_docx(source), audit_docx(output)
    assert before.body_text_sha256 != after.body_text_sha256  # the trap
    assert before.prose_text_sha256 == after.prose_text_sha256

    verified, info = verify_case(
        _case("format_tables", "versioned"),
        {
            "kind": "table_format",
            "tool": "format_tables",
            "source": str(source),
            "output": str(output),
            "table_specs": [
                {
                    "index": 1,
                    "descriptive_columns": [1, 2],
                    "merges": [
                        {"column": 1, "start_row": 2, "end_row": 4, "label": "LoRA"},
                        {"column": 1, "start_row": 5, "end_row": 7, "label": "Full fine-tuning"},
                    ],
                }
            ],
        },
        _transcript("format_tables"),
    )
    assert verified, info
