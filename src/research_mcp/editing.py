"""Versioned prose edits that cannot silently destroy a live citation.

Two probes decided this module's shape before any of it was written.

Assigning to a paragraph's `Range.Text` -- the obvious implementation --
deletes every field in that paragraph with no exception and no warning:
`zotero_item_fields` went 1 -> 0 in the probe. On this project's manuscript
that is 36 live citations one call away from gone. Replacing a *sub-range*
that does not touch a field leaves the field intact and, in the second probe,
also leaves run formatting intact: a bold label beside the edit survived and
the replacement inherited the formatting of the run it landed in.

So every operation here is scoped to a range strictly smaller than the
paragraph, and anything that would intersect a field is refused before Word is
asked to do it.

Addressing is by exact anchor text inside a paragraph whose index came from
`inspect_document`, never by a raw character offset. That keeps offset
arithmetic out of the caller, and out of this project's text extraction: the
one place extraction is known to be unreliable (`text_exact: false`, a
paragraph containing a drawing) is refused outright.
"""

from __future__ import annotations

import hashlib
from pathlib import Path
import tempfile
from typing import Sequence

from .audit import Audit, audit_docx
from .comments import inspect_comments
from .errors import RefusedUnsafe, VerificationFailed
from .paragraphs import ParagraphInfo, read_paragraphs, read_snapshot
from .tables import (
    MergeGroup,
    TableFormatSpec,
    apply_journal_neutral,
    merge_repeated_label,
    snapshot_live,
    verify_saved_journal_neutral,
)
from .word import IsolatedWordDocument, create_versioned_copy


def revision_total(source: str | Path) -> int:
    """All tracked-revision elements, property changes included (plan 501 S2 F2); the audit's
    revision_count sees only ins/del/moves, so it must not gate Word or the direct-XML routing."""

    from .revisions import revision_count

    return revision_count(source)


# Everything an edit must leave alone. `prose_text_sha256` is deliberately
# absent -- changing prose is the point -- so the positive postcondition below
# carries the burden of proving the change was the requested one.
PROTECTED = (
    "table_count",
    "bookmark_names",
    "field_count",
    "zotero_item_fields",
    "zotero_bibliography_fields",
    "revision_count",
    "live_comment_count",
    "drawing_count",
    "live_media_count",
)


def _verify_protected(before: Audit, after: Audit, *, paragraph_delta: int) -> None:
    changed = {
        name: (getattr(before, name), getattr(after, name))
        for name in PROTECTED
        if getattr(before, name) != getattr(after, name)
    }
    expected_paragraphs = before.paragraph_count + paragraph_delta
    if after.paragraph_count != expected_paragraphs:
        changed["paragraph_count"] = (expected_paragraphs, after.paragraph_count)
    if changed:
        raise VerificationFailed(
            "edit changed protected structure", expected="unchanged", actual=changed
        )


def _insert_paragraph_after_body(document, index: int) -> None:
    """Add an empty paragraph after `index` that stays outside any table.

    `Paragraphs(index).Range.InsertParagraphAfter()` inserts at the boundary
    *after* the anchor's paragraph mark. When the next paragraph begins a
    table, Word resolves that boundary into the table's first cell and the new
    paragraph lands inside cell (1,1). Probed on the synthetic fixture:
    `insert_paragraph` reported success while writing its text into the
    existing table, and `Tables.Add` nested a whole table inside that cell.

    Inserting the mark *before* the anchor's own mark splits the anchor
    instead. Both halves are the anchor's own paragraph, so both stay outside
    the table. Word does not merge the resulting adjacent tables; nesting was
    the entire defect.

    The range is collapsed to that point first. On a range that still spans
    the paragraph's content, `InsertParagraphAfter` raises Word's "Cannot edit
    Range" when the content is an inline picture (BIOINF 527 PS2, 2026-11-08:
    a paragraph holding only a drawing). The collapsed range inserts at the
    same position, so prose paragraphs behave as before.
    """

    body = document.Paragraphs(index).Range.Duplicate
    body.End = body.End - 1
    body.Collapse(_WD_COLLAPSE_END)
    body.InsertParagraphAfter()


def _target(paragraphs: tuple[ParagraphInfo, ...], index: int) -> ParagraphInfo:
    if not 1 <= index <= len(paragraphs):
        raise RefusedUnsafe(
            f"paragraph_index {index} is outside 1..{len(paragraphs)}"
        )
    return paragraphs[index - 1]


def _field_spans(document, paragraph_index: int) -> list[tuple[int, int]]:
    """Character spans occupied by fields, in the document's own coordinates.

    A field owns its code and its result. Both are included, because an edit
    landing in either destroys the citation.
    """

    spans: list[tuple[int, int]] = []
    fields = document.Paragraphs(paragraph_index).Range.Fields
    for position in range(1, int(fields.Count) + 1):
        field = fields(position)
        starts, ends = [], []
        for part in ("Code", "Result"):
            try:
                section = getattr(field, part)
            except Exception:  # a field may expose neither on some shapes
                continue
            if section is None:
                continue
            starts.append(int(section.Start))
            ends.append(int(section.End))
        if starts and ends:
            spans.append((min(starts), max(ends)))
    return spans


# Word's Find cannot search for more than 255 characters; a longer find/anchor (up to edit_models.FIND_MAX)
# is found by its first FIND_LIMIT characters and the range is then extended and checked in full (plan 864).
FIND_LIMIT = 255
LONG_FIND_CROSSES_FIELD = "find is {length} characters and crosses a field; use replace_paragraph with paragraph_index"
_WD_FIND_STOP = 0
_WD_COLLAPSE_END = 0
_WD_COLLAPSE_START = 1


def add_comment_to_version(
    source: str | Path, output: str | Path, *, paragraph_index: int,
    anchor: str, text: str, author: str | None = None,
    initials: str | None = None, visible: bool = False,
) -> dict:
    """Add one open comment to a unique visible text anchor without changing prose.

    Direct OOXML (``comment_ops``, plan 562 S1): works on documents with tracked
    changes, keeps every other thread and every revision byte-identical (an
    insertion holding an anchor edge gains only the markers and a run split:
    same author, date, text and formatting), and
    never opens Word (``visible`` is accepted for compatibility and ignored).
    Call through edit_document for source-hash checking, atomic publication
    and undo. Paragraph numbering comes from inspect_document.
    """
    from . import comment_ops

    return comment_ops.add_comment(
        source, output, paragraph_index=paragraph_index, anchor=anchor,
        text=text, author=author, initials=initials,
    )


def _find_in(search, find: str) -> bool:
    finder = search.Find
    finder.ClearFormatting()
    finder.Text = find
    finder.Forward = True
    finder.Wrap = _WD_FIND_STOP
    finder.MatchCase = True
    finder.MatchWildcards = False
    finder.Format = False
    finder.MatchWholeWord = False
    finder.MatchSoundsLike = False
    finder.MatchAllWordForms = False
    return bool(finder.Execute())


def _locate(document, paragraph_index: int, find: str):
    """Ask Word where the anchor is, rather than computing it.

    Range positions and `Range.Text` offsets are different coordinate systems
    once a field is present -- hidden field code occupies positions that the
    text never shows -- so Word does this arithmetic, not this module.

    Word's Find takes at most FIND_LIMIT characters. A longer `find` is
    searched by its first FIND_LIMIT characters; each hit's range is extended
    to the full length and kept only if its `Range.Text` equals `find`, and
    every such match in the paragraph is counted (exactly one is required).
    A prefix hit whose extension reads differently means hidden field code
    sits inside the span: refused with LONG_FIND_CROSSES_FIELD.
    """

    if len(find) <= FIND_LIMIT:
        search = document.Paragraphs(paragraph_index).Range
        if not _find_in(search, find):
            raise RefusedUnsafe(
                f"Word could not find {find!r} in paragraph {paragraph_index}"
            )
        return search
    paragraph = document.Paragraphs(paragraph_index).Range
    start, end = int(paragraph.Start), int(paragraph.End)
    prefix = find[:FIND_LIMIT]
    matches, prefix_hits = [], 0
    while start < end:
        search = document.Range(start, end)
        if not _find_in(search, prefix):
            break
        prefix_hits += 1
        hit = int(search.Start)
        candidate = document.Range(hit, min(hit + len(find), end))
        if str(candidate.Text) == find:
            matches.append(candidate)
        start = hit + 1
    if not matches:
        if prefix_hits:
            raise RefusedUnsafe(LONG_FIND_CROSSES_FIELD.format(length=len(find)))
        raise RefusedUnsafe(
            f"Word could not find the {len(find)}-character text starting {find[:60]!r} in paragraph "
            f"{paragraph_index}"
        )
    if len(matches) > 1:
        raise RefusedUnsafe(
            f"the {len(find)}-character text starting {find[:60]!r} occurs {len(matches)} times in paragraph "
            f"{paragraph_index}; it must occur exactly once"
        )
    return matches[0]


def locate_anchor(document, paragraph_index: int, anchor: str):
    """Resolve `anchor` to a verified Word range inside one paragraph.

    Public so callers outside this module (citation placement) can reuse the
    same field-coordinate-safe lookup instead of computing a position. Refuses
    if Word cannot find it or if the resolved range's text does not match
    `anchor` exactly; an anchor longer than Word's Find limit is found by its
    first FIND_LIMIT characters and checked in full (see `_locate`). Occurrence-count
    uniqueness within the paragraph is the caller's job, checked against the
    paragraph's own text before a Word session exists -- that check needs no
    Word session and should fail before one is spun up.
    """
    candidate = _locate(document, paragraph_index, anchor)
    if str(candidate.Text) != anchor:
        raise RefusedUnsafe(
            f"resolved range does not hold the requested anchor; Word sees "
            f"{str(candidate.Text)!r} where {anchor!r} was expected"
        )
    return candidate


def _refuse_crossing_comment_boundary(document, candidate) -> None:
    for position in range(1, int(document.Comments.Count) + 1):
        scope = document.Comments(position).Scope
        overlaps = int(candidate.Start) < int(scope.End) and int(scope.Start) < int(candidate.End)
        if overlaps and not (
            int(scope.Start) <= int(candidate.Start)
            and int(candidate.End) <= int(scope.End)
        ):
            raise RefusedUnsafe(
                "requested text crosses a comment anchor boundary; make a smaller "
                "edit wholly inside or outside the commented range"
            )


def _format_signature(character) -> tuple:
    font = character.Font
    return (
        int(font.Bold),
        int(font.Italic),
        int(font.Underline),
        int(font.Superscript),
        int(font.Subscript),
        str(font.Name),
        float(font.Size),
    )


def _refuse_unsafe_candidate(document, candidate) -> None:
    if int(candidate.Fields.Count):
        raise RefusedUnsafe("requested text contains a field; edit around it")
    if int(candidate.ContentControls.Count):
        raise RefusedUnsafe("requested text intersects a content control")
    if int(candidate.Hyperlinks.Count):
        raise RefusedUnsafe("requested text intersects a hyperlink")
    if int(candidate.Revisions.Count):
        raise RefusedUnsafe("requested text intersects a tracked revision")
    if int(candidate.InlineShapes.Count):
        raise RefusedUnsafe("requested text intersects an inline drawing")
    _refuse_crossing_comment_boundary(document, candidate)

    signatures = {
        _format_signature(candidate.Characters(position))
        for position in range(1, int(candidate.Characters.Count) + 1)
        if str(candidate.Characters(position).Text) not in ("\r", "\x07")
    }
    if len(signatures) > 1:
        raise RefusedUnsafe(
            "requested text spans mixed formatting; make smaller replacements "
            "inside one uniformly formatted run"
        )


def _refuse_field_overlap(spans, start: int, end: int) -> None:
    for field_start, field_end in spans:
        # Touching at an endpoint is adjacency, not overlap: an anchor that
        # ends exactly where a field begins may be edited safely, and
        # refusing it would block legitimate edits next to every citation.
        if start < field_end and field_start < end:
            raise RefusedUnsafe(
                f"the requested range [{start}, {end}) overlaps a field at "
                f"[{field_start}, {field_end}); editing it would destroy the field"
            )


def _verify_text(
    output_path: Path,
    expected: list[str],
    *,
    label: str,
) -> tuple[ParagraphInfo, ...]:
    saved = read_paragraphs(output_path)
    if len(saved) != len(expected):
        raise VerificationFailed(
            f"{label} changed the paragraph count",
            expected=len(expected),
            actual=len(saved),
        )
    for position, (item, wanted) in enumerate(zip(saved, expected), start=1):
        if item.text != wanted:
            raise VerificationFailed(
                f"{label} produced unexpected text in paragraph {position}",
                expected=wanted,
                actual=item.text,
            )
    return saved


def insert_text_to_version(
    source: str | Path,
    output: str | Path,
    *,
    paragraph_index: int,
    anchor: str,
    position: str,
    text: str,
    visible: bool = False,
) -> dict:
    """Insert text immediately before or after one exact visible anchor."""

    if not anchor:
        raise RefusedUnsafe("anchor must not be empty")
    if not text:
        raise RefusedUnsafe("text must not be empty")
    if position not in ("before", "after"):
        raise RefusedUnsafe("position must be 'before' or 'after'")
    inserted = text + anchor if position == "before" else anchor + text
    if revision_total(source):
        # Plan 864: on a document with tracked revisions (and tracking off for this edit) the insertion is
        # made as untracked direct XML, never through Word, and refused where it touches a revised run.
        from .revisions import plain_replace

        return plain_replace(source, output, op="insert_text", paragraph_index=paragraph_index, find=anchor,
                             replace=inserted)

    before_paragraphs = read_paragraphs(source)
    target = _target(before_paragraphs, paragraph_index)
    if not target.text_exact:
        raise RefusedUnsafe(
            f"paragraph {paragraph_index} contains a drawing, so its text cannot "
            "be matched against Word reliably"
        )
    occurrences = target.text.count(anchor)
    if occurrences != 1:
        raise RefusedUnsafe(
            f"anchor must occur exactly once in paragraph {paragraph_index}; "
            f"it occurs {occurrences} times"
        )

    expected = [item.text for item in before_paragraphs]
    expected[paragraph_index - 1] = target.text.replace(anchor, inserted, 1)

    before = audit_docx(source)
    _, output_path = create_versioned_copy(source, output)
    with IsolatedWordDocument(output_path, visible=visible) as session:
        document = session.document
        candidate = locate_anchor(document, paragraph_index, anchor)
        _refuse_field_overlap(
            _field_spans(document, paragraph_index),
            int(candidate.Start),
            int(candidate.End),
        )
        _refuse_unsafe_candidate(document, candidate)
        insertion = candidate.Duplicate
        insertion.Collapse(
            _WD_COLLAPSE_START if position == "before" else _WD_COLLAPSE_END
        )
        insertion.Text = text
        session.save()

    after = audit_docx(output_path)
    _verify_protected(before, after, paragraph_delta=0)
    _verify_text(output_path, expected, label="insert_text")
    return {
        "source": str(Path(source).resolve()),
        "output": str(output_path),
        "operation": "insert_text",
        "paragraph_index": paragraph_index,
        "position": position,
        "text_before": target.text,
        "text_after": expected[paragraph_index - 1],
        "source_sha256": before.sha256,
        "output_sha256": after.sha256,
        "paragraph_count": after.paragraph_count,
        "indices_after_shifted": False,
        "protected_structure_unchanged": True,
    }


def _table_identity(table) -> tuple:
    return (
        table.rows,
        table.columns,
        table.cell_hashes,
        table.vertical_merges,
        table.horizontal_spans,
    )


def _refuse_table_damage(saved, new_index: int, before: Audit, after: Audit) -> None:
    """Refuse a paragraph insertion that touched a table.

    Split out of `insert_paragraph_to_version` so it can be driven without
    Word. The live tests that exercise it are opt-in, and this guard covers a
    bug that reported success, so it should not be reachable only by a test
    nobody runs by default.
    """

    if saved[new_index - 1].in_table:
        raise VerificationFailed(
            "inserted paragraph landed inside a table",
            expected="body prose",
            actual=f"paragraph {new_index} is in a table",
        )
    if tuple(_table_identity(item) for item in after.tables) != tuple(
        _table_identity(item) for item in before.tables
    ):
        raise VerificationFailed(
            "paragraph insertion changed an existing table",
            expected=tuple(_table_identity(item) for item in before.tables),
            actual=tuple(_table_identity(item) for item in after.tables),
        )


class StyleRef:
    """A paragraph style resolved against ``word/styles.xml`` (plan 864 decision 19).

    ``style_id`` is what ``w:pStyle`` holds; ``name`` is what Word's COM ``Style`` takes. ``defined`` is False for
    a built-in Word knows (a latent style) that the document has not defined yet: Word can apply it by name and
    adds the definition, the direct-XML path cannot."""

    def __init__(self, style_id: str | None, name: str, *, defined: bool = True, default: bool = False) -> None:
        self.style_id, self.name, self.defined, self.default = style_id, name, defined, default

    def __repr__(self) -> str:  # pragma: no cover - debugging aid
        return f"StyleRef({self.style_id!r}, {self.name!r}, defined={self.defined})"


def _style_key(value: str) -> str:
    return "".join(value.split()).casefold()


def _styles_root(source: str | Path):
    import xml.etree.ElementTree as ET
    import zipfile

    with zipfile.ZipFile(source) as archive:
        if "word/styles.xml" not in archive.namelist():
            return None
        return ET.fromstring(archive.read("word/styles.xml"))


def resolve_paragraph_style(source: str | Path, style: str) -> StyleRef:
    """Resolve ``style`` (a styleId or a name, case- and space-insensitive: "Heading 3", "Heading3" and
    "heading 3" are the same) to one paragraph style of the document. Exact matches win over loose ones; two
    styles that both match loosely, or none, refuse naming the near matches."""

    import difflib

    from .ooxml import W

    wanted = (style or "").strip()
    if not wanted:
        raise RefusedUnsafe("style must not be empty; omit it to copy the anchor paragraph's style")
    root = _styles_root(source)
    styles: list[StyleRef] = []
    latent: list[str] = []
    if root is not None:
        for element in root.findall(W + "style"):
            if element.get(W + "type") != "paragraph" or not element.get(W + "styleId"):
                continue
            name_el = element.find(W + "name")
            name = name_el.get(W + "val") if name_el is not None and name_el.get(W + "val") else element.get(W + "styleId")
            styles.append(StyleRef(element.get(W + "styleId"), name,
                                   default=element.get(W + "default") in ("1", "true", "on")))
        latent = [item.get(W + "name") for item in root.iter(W + "lsdException") if item.get(W + "name")]
    for exact in (lambda ref: ref.style_id == wanted, lambda ref: ref.name == wanted):
        found = [ref for ref in styles if exact(ref)]
        if len(found) == 1:
            return found[0]
    key = _style_key(wanted)
    found = [ref for ref in styles if key in (_style_key(ref.style_id), _style_key(ref.name))]
    if len(found) == 1:
        return found[0]
    if found:
        listing = ", ".join(f"{ref.name!r} (styleId {ref.style_id})" for ref in found)
        raise RefusedUnsafe(f"style {wanted!r} matches several paragraph styles: {listing}; pass the exact styleId")
    built_in = [name for name in latent if _style_key(name) == key]
    if built_in:
        return StyleRef(None, built_in[0], defined=False)
    labels = {ref.name: f"{ref.name!r} (styleId {ref.style_id})" for ref in styles}
    labels.update({ref.style_id: f"{ref.name!r} (styleId {ref.style_id})" for ref in styles})
    near = difflib.get_close_matches(wanted, list(labels), n=8, cutoff=0.6) or difflib.get_close_matches(
        key, [_style_key(item) for item in labels], n=8, cutoff=0.6)
    near_labels = list(dict.fromkeys(labels[item] for item in near if item in labels)) or [
        labels[item] for item in labels if _style_key(item) in near][:8]
    hint = f"; near matches: {', '.join(near_labels)}" if near_labels else ""
    raise RefusedUnsafe(f"style {wanted!r} is not a paragraph style of this document{hint}; "
                        "inspect_document lists the styles paragraphs use")


def _style_or_builtin(source: str | Path, style: str):
    """Resolve `style` for a direct-XML insert (plan 973 S3). A style the document defines resolves as before.
    A Word built-in it has not defined yet (Caption, Heading 1-4, Title, List Paragraph, Quote) comes back with
    the definitions a private Word adds for it on a styles-only scratch copy (`builtin_styles.plan`; its refusal,
    such as a clash with an existing style, is passed on); any other undefined style refuses as before.
    Returns (StyleRef, Addition | None)."""

    from . import builtin_styles

    try:
        ref = resolve_paragraph_style(source, style)
    except RefusedUnsafe:
        if builtin_styles.lookup(style) is None:
            raise
        addition = builtin_styles.plan(source, style)
        return StyleRef(addition.style_id, addition.style_name), addition
    if ref.defined:
        return ref, None
    addition = builtin_styles.plan(source, ref.name)
    if addition is None:
        return ref, None
    return StyleRef(addition.style_id, addition.style_name), addition


def _paragraph_style_ids(source: str | Path) -> dict[str, str]:
    """styleId -> w:name of every style in styles.xml (all types)."""

    from .ooxml import W

    root = _styles_root(source)
    if root is None:
        return {}
    found = {}
    for element in root.findall(W + "style"):
        name_el = element.find(W + "name")
        found[element.get(W + "styleId")] = name_el.get(W + "val") if name_el is not None else element.get(
            W + "styleId")
    return found


def _word_style_candidates(ref: StyleRef | None, requested: str) -> list:
    """What Word's COM `Style` is tried with: the name, the styleId, then the built-in's WdBuiltinStyle id (a
    localized Word knows "Caption" only by its id)."""

    from . import builtin_styles

    candidates: list = [ref.name if ref is not None else requested]
    if ref is not None and ref.style_id and ref.style_id not in candidates:
        candidates.append(ref.style_id)
    built_in = builtin_styles.lookup(ref.name if ref is not None else requested)
    if built_in is not None:
        candidates.append(builtin_styles.BUILTINS[built_in])
    return candidates


def _word_style_ref(source: str | Path, style: str) -> StyleRef:
    """Resolve `style` for the Word path: a style the document defines or lists as latent resolves as before;
    a built-in of `builtin_styles.BUILTINS` the document neither defines nor lists (no latentStyles) is still one
    Word adds on first use, so it resolves as an undefined built-in instead of refusing."""

    from . import builtin_styles

    try:
        return resolve_paragraph_style(source, style)
    except RefusedUnsafe:
        built_in = builtin_styles.lookup(style)
        if built_in is None:
            raise
        return StyleRef(None, built_in, defined=False)


def _styles_added(source: str | Path, output: str | Path) -> list[str]:
    """Styles Word's save added to styles.xml (by styleId), named as `builtin_styles` names them."""

    from . import builtin_styles

    before = _paragraph_style_ids(source)
    return [builtin_styles.lookup(name) or name
            for style_id, name in _paragraph_style_ids(output).items() if style_id not in before]


def _is_default_style(source: str | Path, style_id: str | None) -> bool:
    from .ooxml import W

    root = _styles_root(source)
    if root is None or style_id is None:
        return False
    return any(
        element.get(W + "styleId") == style_id and element.get(W + "type") == "paragraph"
        and element.get(W + "default") in ("1", "true", "on")
        for element in root.findall(W + "style"))


def insert_paragraph_to_version(
    source: str | Path,
    output: str | Path,
    *,
    after_paragraph_index: int,
    text: str,
    style: str | None = None,
    tracked: bool | None = None,
    author: str | None = None,
    visible: bool = False,
) -> dict:
    """Insert one new paragraph directly after an existing one.

    Unlike `replace_text` this shifts every later paragraph index by one, so
    indices an agent already holds from `inspect_document` go stale. The
    result says so explicitly.

    Plan 864 S2: `tracked` (the switch, or the op's own override) makes the
    paragraph(s) a tracked insertion by `author`, written as direct XML; a
    document that already has tracked revisions also takes the direct-XML
    path (untracked when `tracked` is false). Both split `text` on "\\n" into
    several paragraphs. `style` is resolved against styles.xml first
    (decision 19), so a name or a styleId both work.
    """

    if not text:
        raise RefusedUnsafe("text must not be empty")

    if tracked or revision_total(source):
        from .builtin_styles import run_with_style
        from .revisions import DEFAULT_AUTHOR, insert_paragraphs

        ref, addition = _style_or_builtin(source, style) if style else (None, None)
        if ref is not None and not ref.defined:
            raise RefusedUnsafe(
                f"style {style!r} is a Word built-in this document has not used yet, so it has no definition in "
                "styles.xml to point the new paragraph at; apply it once in Word (or use a style the document "
                "already has), then insert again")
        # Plan 973 S3: a missing built-in (builtin_styles.BUILTINS) gets the definitions a private Word adds for it
        # on a styles-only scratch copy, spliced into styles.xml (not tracked).
        return run_with_style(source, addition, lambda staged: insert_paragraphs(
            staged, output, after_paragraph_index=after_paragraph_index, text=text,
            style_id=ref.style_id if ref is not None else None,
            default_style=bool(ref is not None and ref.default),
            style_name=ref.name if ref is not None else None,
            tracked=bool(tracked), author=(author or DEFAULT_AUTHOR)))
    ref = _word_style_ref(source, style) if style else None
    if "\n" in text:
        raise RefusedUnsafe(
            "text with \"\\n\" makes several paragraphs, which this document's Word path does not do; "
            "send one insert_paragraph per paragraph (each after the one before), or turn tracking on")

    before_paragraphs = read_paragraphs(source)
    anchor = _target(before_paragraphs, after_paragraph_index)
    if anchor.in_table:
        raise RefusedUnsafe(
            f"paragraph {after_paragraph_index} is inside a table; inserting there "
            "would change table structure, which this operation does not do"
        )

    new_index = after_paragraph_index + 1
    expected = [item.text for item in before_paragraphs]
    expected.insert(after_paragraph_index, text)

    before = audit_docx(source)
    # Decision 19: Word's COM `Style` takes the style's name ("heading 3"),
    # while `w:pStyle` stores its styleId ("Heading3"). The request was
    # resolved once against styles.xml; Word gets the name (then the styleId
    # as a fallback) and the saved paragraph is compared by styleId.
    requested_style = ref.name if ref is not None else "Normal"
    _, output_path = create_versioned_copy(source, output)
    with IsolatedWordDocument(output_path, visible=visible) as session:
        document = session.document
        _insert_paragraph_after_body(document, after_paragraph_index)
        created = document.Paragraphs(new_index)
        # InsertBefore, not InsertAfter: a paragraph's range includes its
        # paragraph mark, so InsertAfter puts the text past the mark and into
        # the *following* paragraph. That produced text in the wrong paragraph
        # until the postcondition caught it.
        created.Range.InsertBefore(text)
        candidates = _word_style_candidates(ref, requested_style)
        last_error: Exception | None = None
        for candidate in candidates:
            try:
                created.Style = candidate
                last_error = None
                break
            except Exception as exc:  # noqa: BLE001 - COM raises bare com_error
                last_error = exc
        if last_error is not None:
            raise RefusedUnsafe(
                f"style {requested_style!r} is not defined in this document"
            ) from last_error
        session.save()

    after = audit_docx(output_path)
    _verify_protected(before, after, paragraph_delta=1)
    saved = _verify_text(output_path, expected, label="insert_paragraph")
    # The anchor is refused inside a table, so a correct insertion is always
    # body prose and every table must survive byte for byte. Without this the
    # old boundary bug passed: the text went into cell (1,1) of the following
    # table, `paragraph_count` still rose by one, `_verify_text` still matched
    # because the corrupted cell held the requested text at the expected
    # index, and the operation reported success. `PROTECTED` counts tables but
    # not their contents, so nothing else here would notice.
    _refuse_table_damage(saved, new_index, before, after)
    saved_style = saved[new_index - 1].style
    # Word omits an explicit pStyle for the document's default paragraph
    # style, so either spelling verifies it; every other style must
    # round-trip by styleId (Word adds a latent built-in's definition on use,
    # so the comparison resolves against the saved document's styles).
    expected_ref = ref if ref is not None and ref.defined else (
        resolve_paragraph_style(output_path, requested_style) if ref is not None else None)
    if expected_ref is None:
        style_matches = saved_style in (None, "Normal") or _is_default_style(output_path, saved_style)
        expected_id = "Normal"
    else:
        expected_id = expected_ref.style_id
        style_matches = saved_style == expected_id or (
            expected_ref.default and saved_style is None)
    if not style_matches:
        raise VerificationFailed(
            "inserted paragraph style did not round-trip",
            expected=expected_id,
            actual=saved_style,
        )
    if expected_ref is not None:
        requested_style = expected_ref.name

    return {
        "source": str(Path(source).resolve()),
        "output": str(output_path),
        "operation": "insert_paragraph",
        "new_paragraph_index": new_index,
        "style": requested_style,
        "style_id": expected_id,
        "styles_added": _styles_added(source, output_path),
        "source_sha256": before.sha256,
        "output_sha256": after.sha256,
        "paragraph_count": after.paragraph_count,
        # Every index at or after this one that the caller already holds is
        # now stale. Two inserts in a row without re-inspecting put the second
        # one in the wrong place.
        "indices_after_shifted": True,
        "stale_from_index": new_index,
        "protected_structure_unchanged": True,
    }


def insert_table_to_version(
    source: str | Path,
    output: str | Path,
    *,
    after_paragraph_index: int,
    headers: Sequence[str],
    rows: Sequence[Sequence[str]],
    descriptive_columns: Sequence[int],
    merges: Sequence[MergeGroup] = (),
    caption: str | None = None,
    caption_style: str | None = None,
    tracked: bool | None = None,
    author: str | None = None,
    visible: bool = False,
) -> dict:
    """Insert one verified journal-neutral table after a prose paragraph.

    Plan 916 S5: `tracked` (the switch, or the op's own override) makes the table
    (and caption) a tracked insertion by `author`, written as direct XML the way
    Word records one (`tracked_table`); a document that already has tracked
    revisions also takes the direct-XML path (untracked when `tracked` is false).
    A revision-free document with tracking off keeps the Word path.
    """

    if not headers or not rows:
        raise RefusedUnsafe("a table requires at least one header and one body row")
    columns = len(headers)
    if any(len(row) != columns for row in rows):
        raise RefusedUnsafe("every body row must have the same width as headers")
    if not descriptive_columns or any(
        column < 1 or column > columns for column in descriptive_columns
    ):
        raise RefusedUnsafe("descriptive columns must be inside the table width")

    if tracked or revision_total(source):
        from . import tracked_table
        from .revisions import DEFAULT_AUTHOR

        from .builtin_styles import run_with_style

        ref = addition = None
        if caption is not None:
            ref, addition = _style_or_builtin(source, caption_style or "Caption")
            if not ref.defined:
                raise RefusedUnsafe(
                    f"caption style {ref.name!r} is a Word built-in this document has not used yet, so it has no "
                    "definition in styles.xml to point the caption at; apply it once in Word (or pass a "
                    "caption_style the document already has), then insert again")
        # Plan 973 S3: a missing Caption (or other built-in of builtin_styles.BUILTINS) gets the definitions a
        # private Word adds for it on a styles-only scratch copy; they are not tracked, so Reject of a tracked
        # insertion leaves an unused style behind, as in Word.
        return run_with_style(source, addition, lambda staged: tracked_table.insert_table(
            staged, output, after_paragraph_index=after_paragraph_index, headers=headers, rows=rows,
            descriptive_columns=descriptive_columns, merges=merges, caption=caption,
            caption_style_id=ref.style_id if ref is not None else None,
            caption_style_name=(caption_style or "Caption") if ref is not None else None,
            tracked=bool(tracked), author=(author or DEFAULT_AUTHOR)))

    before_paragraphs = read_paragraphs(source)
    anchor = _target(before_paragraphs, after_paragraph_index)
    if anchor.in_table:
        raise RefusedUnsafe("a new table cannot be anchored inside an existing table")
    before = audit_docx(source)

    # Plan 973 S3: resolve the caption style against styles.xml first (a name or a styleId, as for
    # insert_paragraph); Word gets its name, then its styleId, then the built-in's WdBuiltinStyle id.
    requested_caption_style = caption_style or "Caption"
    caption_ref = _word_style_ref(source, requested_caption_style) if caption is not None else None
    _, output_path = create_versioned_copy(source, output)
    new_table_index = 0
    with IsolatedWordDocument(output_path, visible=visible) as session:
        document = session.document
        old_live = snapshot_live(document)
        _insert_paragraph_after_body(document, after_paragraph_index)
        insertion_index = after_paragraph_index + 1
        if caption is not None:
            caption_paragraph = document.Paragraphs(insertion_index)
            caption_paragraph.Range.InsertBefore(caption)
            last_error: Exception | None = None
            for candidate in _word_style_candidates(caption_ref, requested_caption_style):
                try:
                    caption_paragraph.Style = candidate
                    last_error = None
                    break
                except Exception as exc:  # noqa: BLE001 - COM raises bare com_error
                    last_error = exc
            if last_error is not None:
                raise RefusedUnsafe(
                    f"caption style {requested_caption_style!r} is not defined"
                ) from last_error
            caption_paragraph.Range.InsertParagraphAfter()
            insertion_index += 1

        table_range = document.Paragraphs(insertion_index).Range.Duplicate
        table_range.Collapse(_WD_COLLAPSE_START)
        table = document.Tables.Add(table_range, len(rows) + 1, columns)
        for row_index, values in enumerate((headers, *rows), start=1):
            for column_index, value in enumerate(values, start=1):
                content = table.Cell(row_index, column_index).Range.Duplicate
                content.End -= 1
                content.Text = value

        table_start = int(table.Range.Start)
        starts = [
            int(document.Tables(index).Range.Start)
            for index in range(1, int(document.Tables.Count) + 1)
        ]
        new_table_index = starts.index(table_start) + 1
        spec = TableFormatSpec(
            index=new_table_index,
            descriptive_columns=tuple(descriptive_columns),
            merges=tuple(merges),
        )
        apply_journal_neutral(document, spec)
        for group in merges:
            merge_repeated_label(table, group)

        all_live = list(snapshot_live(document))
        inserted_live = all_live.pop(new_table_index - 1)
        if tuple(_table_identity(item) for item in old_live) != tuple(
            _table_identity(item) for item in all_live
        ):
            raise VerificationFailed(
                "table insertion changed an existing table",
                expected=tuple(_table_identity(item) for item in old_live),
                actual=tuple(_table_identity(item) for item in all_live),
            )
        if inserted_live.rows != len(rows) + 1 or inserted_live.columns != columns:
            raise VerificationFailed("inserted table has the wrong dimensions")
        session.save()

    after = audit_docx(output_path)
    if after.table_count != before.table_count + 1:
        raise VerificationFailed(
            "table insertion changed the table count unexpectedly",
            expected=before.table_count + 1,
            actual=after.table_count,
        )
    protected = (
        "bookmark_names",
        "field_count",
        "zotero_item_fields",
        "zotero_bibliography_fields",
        "revision_count",
        "live_comment_count",
        "drawing_count",
        "live_media_count",
    )
    changed = {
        name: (getattr(before, name), getattr(after, name))
        for name in protected
        if getattr(before, name) != getattr(after, name)
    }
    if changed:
        raise VerificationFailed(
            "table insertion changed protected structure",
            expected="unchanged",
            actual=changed,
        )
    old_saved = list(after.tables)
    inserted = old_saved.pop(new_table_index - 1)
    if tuple(_table_identity(item) for item in before.tables) != tuple(
        _table_identity(item) for item in old_saved
    ):
        raise VerificationFailed("saved table insertion changed an existing table")
    expected_grid = [list(headers), *[list(row) for row in rows]]
    expected_hashes = tuple(
        hashlib.sha256(" ".join(value.split()).encode("utf-8")).hexdigest()
        for grid_row in expected_grid
        for value in grid_row
    )
    if inserted.cell_hashes != expected_hashes:
        raise VerificationFailed(
            "inserted table cell text did not round-trip",
            expected=expected_hashes,
            actual=inserted.cell_hashes,
        )
    expected_merges = tuple(
        (group.start_row, group.end_row, group.column, group.label)
        for group in merges
    )
    if inserted.vertical_merges != expected_merges:
        raise VerificationFailed(
            "inserted table merges did not round-trip",
            expected=expected_merges,
            actual=inserted.vertical_merges,
        )
    verify_saved_journal_neutral(
        output_path,
        (TableFormatSpec(new_table_index, tuple(descriptive_columns), tuple(merges)),),
    )

    if caption is not None:
        matches = [item for item in read_paragraphs(output_path) if item.text == caption]
        if len(matches) != 1:
            raise VerificationFailed(
                "table caption did not round-trip exactly",
                expected=caption,
                actual=[item.text for item in matches],
            )
        # w:pStyle holds the styleId, and Word omits it for the default paragraph style (Normal), so the saved
        # caption is compared by the styleId the request resolves to in the saved document; the default style
        # also matches an absent pStyle. Comparing the requested name itself refused caption_style="Normal".
        expected_ref = caption_ref if caption_ref.defined else resolve_paragraph_style(
            output_path, caption_ref.name)
        saved_style = matches[0].style
        if not (saved_style == expected_ref.style_id or (expected_ref.default and saved_style is None)):
            raise VerificationFailed(
                "table caption style did not round-trip",
                expected=expected_ref.style_id,
                actual=saved_style,
            )

    return {
        "source": str(Path(source).resolve()),
        "output": str(output_path),
        "operation": "insert_table",
        "table_index": new_table_index,
        "rows": inserted.rows,
        "columns": inserted.columns,
        "caption": caption,
        "caption_style": expected_ref.name if caption is not None else None,
        "styles_added": _styles_added(source, output_path),
        "source_sha256": before.sha256,
        "output_sha256": after.sha256,
        "paragraph_count": after.paragraph_count,
        "indices_after_shifted": True,
        "stale_from_index": after_paragraph_index + 1,
        "protected_structure_unchanged": True,
    }


from .comment_ops import DISPOSITION_LABELS as _DISPOSITION_LABELS  # noqa: E402


def reply_comment_to_version(
    source: str | Path,
    output: str | Path,
    *,
    comment_id: str,
    disposition: str,
    reply_text: str,
    resolve: bool = False,
    author: str | None = None,
    initials: str | None = None,
    visible: bool = False,
) -> dict:
    """Reply to one comment thread; it stays open unless ``resolve`` is true.

    Direct OOXML (``comment_ops``, plan 562 S1); ``visible`` is ignored.
    """
    from . import comment_ops

    return comment_ops.reply_comment(
        source, output, comment_id=comment_id, disposition=disposition,
        reply_text=reply_text, resolve=resolve, author=author, initials=initials,
    )


# --------------------------------------------------------------------------
# Minimal-diff text edits (plan 328, D8/D9/D11/D12).
#
# `sections.py` plans which characters, paragraphs and rows change; this part
# performs the plan in Word, bottom-up, and the saved candidate is then held
# to the plan by `sections.verify_plan`. No unchanged character is ever
# retyped: new text is inserted at a collapsed range (`_insert`) carrying
# the formatting of the character the plan names, and replaced characters
# are deleted around it, so every other run, bookmark and comment boundary
# is left where it was and Word's own typing rule places the new text
# (`sections` documents it; `verify_plan` holds every bookmark and comment
# extent to it).
# --------------------------------------------------------------------------

_WD_FIELD_HYPERLINK = 88
_SKIP_CHARACTERS = ("", "\x13", "\x14", "\x15")


def _short_text(text: str, limit: int = 40) -> str:
    text = text.replace("\r", " ").replace("\x07", " ").strip()
    return text if len(text) <= limit else text[: limit - 1] + "…"


def _field_result_spot(document, paragraph, low: int, high: int, wanted: str, flags, after: int) -> list[int]:
    """The position in a field result inside [low, high) holding `wanted`."""

    from .sections import caps_equal

    fields = paragraph.Fields
    for number in range(1, int(fields.Count) + 1):
        result = fields(number).Result
        start, end = int(result.Start), int(result.End)
        if not (low <= start and end <= high):
            continue
        for spot in range(max(start, after + 1), end):
            text = str(document.Range(spot, spot + 1).Text)
            if len(text) == 1 and caps_equal(wanted, flags, text):
                return [spot]
    return []


def _char_positions(document, index: int, stored: str, caps, op: str) -> list[int]:
    """Word position of every stored character of paragraph `index`."""

    from .sections import caps_equal

    paragraph = document.Paragraphs(index).Range
    start = int(paragraph.Start)
    width = len(stored)
    simple = not (
        int(paragraph.Fields.Count) or int(paragraph.ContentControls.Count) or int(paragraph.InlineShapes.Count)
    )
    if simple:
        visible = str(document.Range(start, start + width).Text)
        mark = str(document.Range(start + width, start + width + 1).Text)
        if caps_equal(stored, caps, visible) and mark in ("\r", "\x07") and int(paragraph.End) == start + width + 1:
            return list(range(start, start + width))
    positions: list[int] = []
    for character in paragraph.Characters:
        text = str(character.Text)
        if text in _SKIP_CHARACTERS:
            continue
        if len(positions) == width and text in ("\r", "\x07", "\r\x07"):
            break
        if len(positions) < width and len(text) == 1 and caps_equal(stored[len(positions)], caps[len(positions):len(positions) + 1], text):
            # At a field boundary Word widens the character: the first
            # character of a hyperlink result reported Start 717 (the begin
            # marker) for a 't' at 741, and the last character of a field
            # result reports its End past the end marker. Pick the one
            # position inside the span that holds the character.
            low, high = int(character.Start), int(character.End)
            position = low
            if high - low != 1:
                wanted = stored[len(positions)]
                flags = caps[len(positions):len(positions) + 1]
                # End - 1 first (begin-marker case), then Start (end-marker
                # case); a field's code may repeat the character elsewhere.
                hits = [
                    spot for spot in (high - 1, low)
                    if len(str(document.Range(spot, spot + 1).Text)) == 1
                    and caps_equal(wanted, flags, str(document.Range(spot, spot + 1).Text))
                ][:1]
                if not hits:
                    # A field whose result is one character (a citation
                    # superscript "1") widens at both ends at once: Word
                    # reported [366, 2317) for the '1' at 2315, i.e. begin
                    # marker to past the end marker. Look inside the result.
                    hits = _field_result_spot(
                        document, paragraph, low, high, wanted, flags,
                        positions[-1] if positions else -1)
                if not hits:
                    raise RefusedUnsafe(
                        f"{op}: cannot place character {len(positions)} of paragraph {index} in Word "
                        f"(span [{low}, {high})); inspect the document again")
                position = hits[0]
            positions.append(position)
            continue
        raise RefusedUnsafe(
            f"{op}: paragraph {index} in Word does not match the saved text at character {len(positions)} "
            f"(Word shows {text!r}); inspect the document again"
        )
    if len(positions) != width:
        raise RefusedUnsafe(
            f"{op}: paragraph {index} in Word has {len(positions)} visible characters; the saved text has {width}"
        )
    return positions


def _paragraph_objects(document, index: int) -> list[dict]:
    """Fields, hyperlinks, content controls, drawings and comments touching a paragraph."""

    paragraph = document.Paragraphs(index).Range
    start, end = int(paragraph.Start), int(paragraph.End)
    objects: list[dict] = []
    fields = paragraph.Fields
    for position in range(1, int(fields.Count) + 1):
        field = fields(position)
        try:
            code = field.Code
            result = field.Result
            outer = (int(code.Start) - 1, int(result.End) + 1)
            interior = (int(result.Start), int(result.End))
            code_text = str(code.Text).strip()
            shown = str(result.Text)
        except Exception:
            continue
        if int(field.Type) == _WD_FIELD_HYPERLINK:
            objects.append({"kind": "hyperlink", "outer": outer, "interior": interior,
                            "name": f"hyperlink {_short_text(shown)!r}"})
        else:
            label = code_text.split(" ", 2)
            label = " ".join(label[:2]) if label and label[0].upper() == "ADDIN" else (label[0] if label else "")
            objects.append({"kind": "field", "outer": outer, "interior": None,
                            "name": f"field {label} showing {_short_text(shown)!r}"})
    controls = paragraph.ContentControls
    for position in range(1, int(controls.Count) + 1):
        control = controls(position)
        inner = control.Range
        name = str(control.Title or control.Tag or "") or _short_text(str(inner.Text))
        objects.append({"kind": "content control", "outer": (int(inner.Start) - 1, int(inner.End) + 1),
                        "interior": (int(inner.Start), int(inner.End)), "name": f"content control {name!r}"})
    shapes = paragraph.InlineShapes
    for position in range(1, int(shapes.Count) + 1):
        shape = shapes(position).Range
        objects.append({"kind": "drawing", "outer": (int(shape.Start), int(shape.End)), "interior": None,
                        "name": "an inline drawing"})
    for position in range(1, int(document.Comments.Count) + 1):
        comment = document.Comments(position)
        scope = comment.Scope
        a, b = int(scope.Start), int(scope.End)
        if a < end and start < b:
            objects.append({"kind": "comment", "outer": (a, b), "interior": (a, b),
                            "name": f"comment {position} ({str(comment.Author)}) on {_short_text(str(scope.Text))!r}"})
    return objects


def _refuse_cut(objects: list[dict], span: tuple[int, int], what: str) -> None:
    a, b = span
    for item in objects:
        o1, o2 = item["outer"]
        if not (a < o2 and o1 < b):
            continue
        interior = item["interior"]
        if interior is not None and interior[0] <= a and b <= interior[1]:
            continue  # wholly inside the visible text of a hyperlink, content control or comment scope
        if item["kind"] == "field":
            raise RefusedUnsafe(
                f"{what} overlaps a field: it would cut through {item['name']} "
                f"(range [{a}, {b}) vs field [{o1}, {o2})); editing it would destroy the field")
        raise RefusedUnsafe(f"{what} would cut through {item['name']}; make the edit wholly inside or outside it")


def _insert(document, position: int, text: str, like: int, paragraph_start: int) -> None:
    """Insert `text` at the collapsed `position`, formatted like the character at `like`.

    Word gives text inserted at a collapsed range the formatting of the
    character before it (at a paragraph's start, the one after), so that is
    used as is when it is the wanted one. Otherwise (`like == position`) a
    formatted copy of that character goes in first, the text is inserted
    after the copy, and the copy is removed: the text gets the character's
    exact run properties, and lands on the same side of every bookmark and
    comment boundary as a plain insertion at `position` would.
    """

    if like == position - 1 or position == paragraph_start:
        document.Range(position, position).InsertAfter(text)
        return
    if like != position:  # pragma: no cover - callers pass a neighbour
        raise RefusedUnsafe(f"internal: formatting source {like} is not next to insertion point {position}")
    document.Range(position, position).FormattedText = document.Range(like, like + 1).FormattedText
    document.Range(position + 1, position + 1).InsertAfter(text)
    document.Range(position, position + 1).Text = ""


def _anchor_edges(document, start: int, end: int) -> set[int]:
    """Word positions in [start, end] where a bookmark starts or ends, or a comment's scope ends."""

    edges: set[int] = set()
    bookmarks = document.Bookmarks
    hidden = bool(bookmarks.ShowHidden)
    try:
        bookmarks.ShowHidden = True
        for position in range(1, int(bookmarks.Count) + 1):
            extent = bookmarks(position).Range
            edges.update(spot for spot in (int(extent.Start), int(extent.End)) if start <= spot <= end)
    finally:
        bookmarks.ShowHidden = hidden
    for position in range(1, int(document.Comments.Count) + 1):
        spot = int(document.Comments(position).Scope.End)
        if start <= spot <= end:
            edges.add(spot)
    return edges


def _replace(document, first: int, last: int, text: str, paragraph_start: int, edges) -> None:
    """Replace Word characters [first, last) with `text`, formatted like the character at `first`.

    The new text goes in strictly inside the replaced characters (after the
    first), so it belongs to every bookmark and comment that held them all;
    then the old characters around it are deleted. A single character has
    no inside: when a bookmark or comment boundary sits right after it, the
    text goes in before it instead.
    """

    width = len(text)
    if last - first == 1 and last in edges():
        _insert(document, first, text, first, paragraph_start)
        document.Range(first + width, first + width + 1).Text = ""
        return
    document.Range(first + 1, first + 1).InsertAfter(text)
    if last - first > 1:
        document.Range(first + 1 + width, last + width).Text = ""
    document.Range(first, first + 1).Text = ""


def _text_step(document, step, op: str) -> None:
    positions = _char_positions(document, step.index, step.text, step.caps, op)
    objects = _paragraph_objects(document, step.index)
    paragraph = document.Paragraphs(step.index).Range
    paragraph_start = int(paragraph.Start)
    found: list[set[int]] = []

    def edges() -> set[int]:
        if not found:
            found.append(_anchor_edges(document, paragraph_start, int(paragraph.End)))
        return found[0]

    for edit in reversed(step.edits):
        what = f"{op}: {edit.describe(step.text)} of paragraph {step.index}"
        if edit.mode == "empty":
            document.Range(paragraph_start, paragraph_start).InsertBefore(edit.text)
            continue
        if edit.mode in ("replace", "delete"):
            first, last = positions[edit.start], positions[edit.end - 1] + 1
            if last - first != edit.end - edit.start:
                raise RefusedUnsafe(f"{what} would cut through hidden content (a field code or control marker)")
            _refuse_cut(objects, (first, last), what)
            if edit.mode == "delete":
                document.Range(first, last).Text = ""
            else:
                _replace(document, first, last, edit.text, paragraph_start, edges)
            continue
        like = positions[edit.source]
        point = like + 1 if edit.mode == "after" else like
        _refuse_cut(objects, (point, point), what)
        _insert(document, point, edit.text, like, paragraph_start)


def _check_paragraph_text(document, index: int, stored: str, caps, op: str) -> None:
    from .sections import caps_equal

    shown = str(document.Paragraphs(index).Range.Text).rstrip("\r\x07")
    if not caps_equal(stored, caps, shown):
        raise RefusedUnsafe(
            f"{op}: paragraph {index} in Word reads {shown!r}, not the saved {stored!r}; inspect the document again")


def _execute_plan(document, plan, *, paragraph_count: int) -> None:
    from .sections import DeleteParasStep, DeleteRowStep, DuplicateRowStep, InsertLinesStep, TextStep

    op = plan.op
    structural = any(not isinstance(step, TextStep) for step in plan.steps)
    if structural and int(document.Paragraphs.Count) != paragraph_count:
        raise RefusedUnsafe(
            f"{op}: Word counts {int(document.Paragraphs.Count)} paragraphs where the saved file has "
            f"{paragraph_count}; row and section edits need the two to agree")
    for step in plan.steps:
        if isinstance(step, TextStep):
            _text_step(document, step, op)
        elif isinstance(step, InsertLinesStep):
            joined = "\r".join(step.lines)
            if not step.text:
                start = int(document.Paragraphs(step.index).Range.Start)
                document.Range(start, start).InsertBefore(joined + "\r" if step.where == "before" else "\r" + joined)
                continue
            positions = _char_positions(document, step.index, step.text, step.caps, op)
            objects = _paragraph_objects(document, step.index)
            what = f"{op}: insert {len(step.lines)} paragraph(s) {step.where} paragraph {step.index}"
            if step.where == "before":
                like, point, text = positions[0], positions[0], joined + "\r"
            else:
                like, point, text = positions[-1], positions[-1] + 1, "\r" + joined
            _refuse_cut([dict(item, interior=None) for item in objects], (point, point), what)
            _insert(document, point, text, like, int(document.Paragraphs(step.index).Range.Start))
        elif isinstance(step, DeleteParasStep):
            for index, (text, caps) in enumerate(zip(step.texts, step.caps), start=step.first):
                _check_paragraph_text(document, index, text, caps, op)
            if not step.trailing:
                document.Range(
                    int(document.Paragraphs(step.first).Range.Start),
                    int(document.Paragraphs(step.last).Range.End),
                ).Text = ""
                continue
            previous = document.Paragraphs(step.previous)
            saved = previous.Format.Duplicate if step.fix_format else None
            document.Range(
                int(previous.Range.End) - 1, int(document.Paragraphs(step.last).Range.End) - 1
            ).Text = ""
            if saved is not None:
                document.Paragraphs(step.previous).Format = saved
        elif isinstance(step, DuplicateRowStep):
            template = document.Paragraphs(step.template).Range.Rows(1)
            formatted = template.Range.FormattedText
            if step.before is not None:
                position = int(document.Paragraphs(step.before).Range.Start)
            else:
                position = int(document.Paragraphs(step.append_after).Range.Rows(1).Range.End)
            document.Range(position, position).FormattedText = formatted
        elif isinstance(step, DeleteRowStep):
            document.Paragraphs(step.index).Range.Rows(1).Delete()
        else:  # pragma: no cover - planner and executor are versioned together
            raise RefusedUnsafe(f"{op}: unknown plan step {type(step).__name__}")


def _replan_on_baseline(op: str, make_plan, plan, layout, baseline: Path):
    """The same edit planned on Word's own save of the untouched document.

    The baseline becomes the reference every unchanged character and unit
    is verified against, so it is adopted only after it is shown to preserve
    the source: text, every character's formatting, paragraph/table/row/cell
    properties, bookmark and comment extents, field instructions, content
    controls and hyperlinks (`sections.baseline_preservation_errors`; ids
    may be renumbered and runs re-split). Otherwise nothing is edited.
    """

    from .sections import baseline_preservation_errors, read_layout

    baseline_layout = read_layout(baseline)
    shape = lambda item: [(type(step).__name__, getattr(step, "index", None)) for step in item.steps]  # noqa: E731
    texts = lambda lay: [para.text for para in lay.paragraphs]  # noqa: E731
    if texts(baseline_layout) != texts(layout):
        raise VerificationFailed(
            f"{op}: Word changed the document's text merely by opening and saving it "
            "(for example by updating fields); nothing was edited. Open and save it in Word, then inspect again.")
    problems = baseline_preservation_errors(layout.path, baseline)
    if problems:
        raise VerificationFailed(
            f"{op}: Word's own save of the untouched document did not preserve it, so it cannot serve as the "
            "reference for verifying this edit; nothing was edited: " + "; ".join(problems[:5]),
            expected="source preserved", actual=problems[:20])
    rebased = make_plan(baseline_layout)
    if shape(rebased) != shape(plan):
        raise VerificationFailed(f"{op}: the edit plans differently on Word's own save of the document; nothing was edited")
    return rebased


def _run_plan(source, output, op: str, make_plan, *, visible: bool = False, timeout_seconds: float = 90.0) -> dict:
    from .sections import DeleteRowStep, DuplicateRowStep, read_layout, verify_plan

    if revision_total(source):  # before planning: a revised document never reaches the Word planner
        raise RefusedUnsafe(
            f"{op}: source contains {revision_total(source)} tracked revision elements; {op} runs through Word, "
            "and any Word save rewrites other authors' revision records, so it is refused here. On a document "
            "with revisions use replace_text/replace_paragraph/insert_text (direct XML; tracked when "
            "set_track_changes is on), or accept_revisions/reject_revisions"
        )
    layout = read_layout(source)
    plan = make_plan(layout)
    before = audit_docx(source)
    _, output_path = create_versioned_copy(source, output)
    reference: str | Path = source
    with tempfile.TemporaryDirectory(prefix="research-mcp-baseline-") as scratch:
        if plan.steps:
            with IsolatedWordDocument(output_path, visible=visible, timeout_seconds=timeout_seconds) as session:
                # Word rewrites some XML on any save (bookmark/comment ids, drawing edit ids, empty
                # sdtEndPr). Saving the untouched document first gives a reference in Word's own form,
                # so "everything outside the edit is unchanged" is checked exactly, not loosened.
                baseline = session.save_as(Path(scratch) / "baseline.docx")
                plan = _replan_on_baseline(op, make_plan, plan, layout, baseline)
                reference = baseline
                _execute_plan(session.document, plan, paragraph_count=len(read_layout(baseline).paragraphs))
                session.save_as(output_path)
        after = audit_docx(output_path)
        row_delta = sum(isinstance(step, DuplicateRowStep) for step in plan.steps) - sum(
            isinstance(step, DeleteRowStep) for step in plan.steps)
        # plan.paragraph_delta counts layout indices, which include one row-end marker per table
        # row; the audit counts w:p elements only.
        # verify_plan first: it names a bookmark/comment that an edit would delete, where the
        # protected-structure audit only reports changed counts.
        verification = verify_plan(reference, output_path, plan)
        _verify_protected(before, after, paragraph_delta=plan.paragraph_delta - row_delta)
    verification["compared_against"] = "word_baseline" if reference != source else "source"
    return {
        "source": str(Path(source).resolve()),
        "output": str(output_path),
        "operation": op,
        **plan.report,
        "source_sha256": before.sha256,
        "output_sha256": after.sha256,
        "paragraph_count": after.paragraph_count,
        "paragraph_delta": plan.paragraph_delta - row_delta,
        "index_delta": plan.paragraph_delta,
        "indices_after_shifted": plan.paragraph_delta != 0,
        "shift_base": plan.last_index + 1,
        "stale_range": [plan.first_index, plan.last_index] if plan.paragraph_delta else None,
        "protected_structure_unchanged": True,
        "verification": verification,
    }


def replace_text_to_version(
    source: str | Path,
    output: str | Path,
    *,
    paragraph_index: int,
    find: str,
    replace: str,
    visible: bool = False,
) -> dict:
    """Replace one occurrence of `find` inside one paragraph, changing only differing characters.

    `find` may span runs with different formatting (a whole résumé title
    line). Characters common to `find` and `replace` keep their runs; new text
    takes the formatting of the character it replaces or extends (D8). Case
    is relaxed only on characters inside caps/small-caps runs (D12). Besides
    a missing or repeated `find`, refused when a changed character would cut
    through a field, hyperlink, content control, drawing or comment boundary
    (naming the edit and object), when the paragraph holds a drawing or sits
    in a nested table, or when Word's save of the untouched document does not
    preserve it.

    On a document with tracked revisions the edit is made as direct XML
    (plan 501, `revisions.plain_replace`), never through Word, and refused
    when it would touch a revised run (the refusal names the revision's id,
    type and author); every revision stays byte-identical.
    """

    from .sections import plan_replace_text

    if not find:
        raise RefusedUnsafe("find must not be empty")
    # No FIND_LIMIT here: the plan matches `find` in the layout's text, not with Word's Find (plan 864).
    if revision_total(source):
        from .revisions import plain_replace

        return plain_replace(source, output, op="replace_text", paragraph_index=paragraph_index, find=find,
                             replace=replace)
    return _run_plan(
        source, output, "replace_text",
        lambda layout: plan_replace_text(layout, "replace_text", paragraph_index, find, replace),
        visible=visible,
    )


def replace_paragraph_to_version(
    source: str | Path, output: str | Path, *, paragraph_index: int, text: str, visible: bool = False,
) -> dict:
    """Make paragraph `paragraph_index` read exactly `text`, by minimal character diff (D8).

    On a document with tracked revisions: direct XML, refused where a revised run is touched (see
    replace_text_to_version).
    """

    from .sections import plan_paragraph

    if revision_total(source):
        from .revisions import plain_replace

        return plain_replace(source, output, op="replace_paragraph", paragraph_index=paragraph_index, text=text)
    return _run_plan(
        source, output, "replace_paragraph",
        lambda layout: plan_paragraph(layout, "replace_paragraph", paragraph_index, text),
        visible=visible,
    )


def replace_table_row_to_version(
    source: str | Path, output: str | Path, *, match: str, cells: Sequence[str],
    table_index: int | None = None, tracked: bool | None = None, author: str | None = None,
    visible: bool = False,
) -> dict:
    """Rewrite one table row, found by one cell's (or one cell line's) text (D9).

    `cells` gives every cell's new text; "\\n" separates paragraphs in a
    cell. The cell count must stay the same; paragraphs inside a cell may be
    added or removed.

    Plan 864 S3a: `tracked` (the switch, or the op's own override) records the
    row's changed lines, added lines and removed lines as tracked changes by
    `author`, as direct XML (`tracked_sections`); a document that already has
    tracked revisions takes the same path untracked when `tracked` is false.
    """

    from .sections import plan_table_row

    make_plan = lambda layout: plan_table_row(  # noqa: E731
        layout, "replace_table_row", match, list(cells), table_index=table_index)
    if tracked or revision_total(source):
        from .revisions import DEFAULT_AUTHOR
        from .sections import read_layout
        from .tracked_sections import run_plan_direct

        return run_plan_direct(source, output, make_plan(read_layout(source)), tracked=bool(tracked),
                               author=author or DEFAULT_AUTHOR)
    return _run_plan(source, output, "replace_table_row", make_plan, visible=visible)


def replace_section_to_version(
    source: str | Path, output: str | Path, *, heading: str,
    paragraphs: Sequence[str] | None = None, rows: Sequence[Sequence[str]] | None = None,
    until_heading: str | None = None, tracked: bool | None = None, author: str | None = None,
    visible: bool = False,
) -> dict:
    """Rewrite everything between `heading` and the next heading of its kind (D9).

    A section laid out as table rows takes `rows` (cell texts per row); one
    of body paragraphs takes `paragraphs`. Rows/paragraphs may be added or
    removed only inside the section; everything outside is verified unchanged.

    Plan 864 S3a: as `replace_table_row`, `tracked` records the changed,
    added and removed paragraphs (or row lines) as tracked changes by
    `author`, as direct XML; a revised document takes the same path untracked
    when `tracked` is false. Plan 864 S3b: a table row added or removed runs
    there too, tracked as Word records a row insertion or deletion.
    """

    from .sections import plan_section

    make_plan = lambda layout: plan_section(  # noqa: E731
        layout, "replace_section", heading,
        paragraphs=list(paragraphs) if paragraphs is not None else None,
        rows=[list(row) for row in rows] if rows is not None else None,
        until_heading=until_heading,
    )
    if tracked or revision_total(source):
        from .revisions import DEFAULT_AUTHOR
        from .sections import read_layout
        from .tracked_sections import run_plan_direct

        return run_plan_direct(source, output, make_plan(read_layout(source)), tracked=bool(tracked),
                               author=author or DEFAULT_AUTHOR)
    return _run_plan(source, output, "replace_section", make_plan, visible=visible)
