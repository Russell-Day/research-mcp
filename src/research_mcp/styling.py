"""Academic journal manuscript typography and layout policy enforcement."""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any, Literal
import zipfile
import xml.etree.ElementTree as ET

from .audit import Audit, audit_docx
from .errors import RefusedUnsafe, VerificationFailed
from .word import IsolatedWordDocument, staged_document

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


@dataclass(frozen=True)
class JournalStyleProfile:
    name: str
    font_name: str = "Times New Roman"
    font_size_pt: float = 11.0
    title_size_pt: float = 14.0
    heading_size_pt: float = 11.0
    line_spacing_rule: int = 2  # wdLineSpaceDouble = 2, wdLineSpaceSingle = 0
    line_spacing_pt: float = 24.0
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0
    title_space_before_pt: float = 12.0
    title_space_after_pt: float = 12.0
    heading_space_before_pt: float = 12.0
    heading_space_after_pt: float = 6.0
    alignment: int = 0  # wdAlignParagraphLeft = 0
    margin_top_inches: float = 1.0
    margin_bottom_inches: float = 1.0
    margin_left_inches: float = 1.0
    margin_right_inches: float = 1.0
    strip_header_footers: bool = True  # Double-anonymized review compliance


PROFILES: dict[str, JournalStyleProfile] = {
    "radiology_ai": JournalStyleProfile(
        name="radiology_ai",
        font_name="Times New Roman",
        font_size_pt=11.0,
        title_size_pt=14.0,
        heading_size_pt=11.0,
        line_spacing_rule=2,
        line_spacing_pt=24.0,
        space_before_pt=0.0,
        space_after_pt=0.0,
        title_space_before_pt=12.0,
        title_space_after_pt=12.0,
        heading_space_before_pt=12.0,
        heading_space_after_pt=6.0,
        alignment=0,
        margin_top_inches=1.0,
        margin_bottom_inches=1.0,
        margin_left_inches=1.0,
        margin_right_inches=1.0,
        strip_header_footers=True,
    ),
    "editing_compact": JournalStyleProfile(
        name="editing_compact",
        font_name="Times New Roman",
        font_size_pt=11.0,
        title_size_pt=14.0,
        heading_size_pt=11.0,
        line_spacing_rule=0,  # wdLineSpaceSingle = 0
        line_spacing_pt=12.0,
        space_before_pt=0.0,
        space_after_pt=0.0,
        title_space_before_pt=6.0,
        title_space_after_pt=6.0,
        heading_space_before_pt=6.0,
        heading_space_after_pt=2.0,
        alignment=0,
        margin_top_inches=1.0,
        margin_bottom_inches=1.0,
        margin_left_inches=1.0,
        margin_right_inches=1.0,
        strip_header_footers=True,
    ),
    "compact_single": JournalStyleProfile(
        name="compact_single",
        font_name="Times New Roman",
        font_size_pt=11.0,
        title_size_pt=14.0,
        heading_size_pt=11.0,
        line_spacing_rule=0,  # wdLineSpaceSingle = 0
        line_spacing_pt=12.0,
        space_before_pt=0.0,
        space_after_pt=0.0,
        title_space_before_pt=6.0,
        title_space_after_pt=6.0,
        heading_space_before_pt=6.0,
        heading_space_after_pt=2.0,
        alignment=0,
        margin_top_inches=1.0,
        margin_bottom_inches=1.0,
        margin_left_inches=1.0,
        margin_right_inches=1.0,
        strip_header_footers=True,
    ),
    "standard_double": JournalStyleProfile(
        name="standard_double",
        font_name="Times New Roman",
        font_size_pt=12.0,
        title_size_pt=14.0,
        heading_size_pt=12.0,
        line_spacing_rule=2,
        line_spacing_pt=24.0,
        space_before_pt=0.0,
        space_after_pt=0.0,
        title_space_before_pt=12.0,
        title_space_after_pt=12.0,
        heading_space_before_pt=12.0,
        heading_space_after_pt=6.0,
        alignment=0,
        margin_top_inches=1.0,
        margin_bottom_inches=1.0,
        margin_left_inches=1.0,
        margin_right_inches=1.0,
        strip_header_footers=False,
    ),
}

MAJOR_HEADINGS = {
    "introduction",
    "materials and methods",
    "methods",
    "results",
    "discussion",
    "conclusion",
    "references",
    "structured abstract",
    "key points",
}


def classify_paragraph(text: str, style_name: str, is_first_body: bool, in_table: bool) -> str:
    """One explicit role rule shared by planning, application and verification."""
    if in_table:
        return "table"
    if is_first_body or "title" in style_name.lower():
        return "title"
    if text.strip().lower().rstrip(":") in MAJOR_HEADINGS or style_name.lower().startswith("heading"):
        return "heading"
    if text.strip().startswith("Summary Statement:"):
        return "summary"
    return "body"


def _style_name(paragraph: Any) -> str:
    """Return a paragraph style name without letting COM quirks abort styling."""
    try:
        return str(paragraph.Style.NameLocal or "")
    except Exception:
        return ""


def _apply_line_spacing(fmt: Any, profile: JournalStyleProfile) -> None:
    """Set Word's spacing *rule* without a conflicting LineSpacing assignment.

    Word can silently turn a single/double rule into an exact or multiple rule
    when ``LineSpacing`` is assigned afterwards.  The named Word rules already
    carry the correct line pitch, so setting the numeric property is both
    unnecessary and a source of intermittent layout differences.
    """
    fmt.LineSpacingRule = profile.line_spacing_rule
    if profile.line_spacing_rule not in (0, 2):
        fmt.LineSpacing = profile.line_spacing_pt


def _strip_headers_footers(doc: Any) -> None:
    """Clear existing headers and footers, surfacing COM failures.

    Header/footer access can fail for protected or malformed sections.  An
    anonymized output must never report success while an identifying header or
    footer was left behind, so errors are collected and raised after all
    sections have been attempted.
    """
    failures: list[str] = []
    for section_number, section in enumerate(doc.Sections, start=1):
        for kind, collection in (("header", section.Headers), ("footer", section.Footers)):
            for index in (1, 2, 3):  # Primary, FirstPage, EvenPages
                try:
                    item = collection(index)
                    if not bool(item.Exists):
                        continue
                    item.Range.Text = ""
                except Exception as exc:
                    failures.append(f"section {section_number} {kind} {index}: {exc}")
    if failures:
        raise VerificationFailed(
            "could not clear all headers/footers required for anonymized styling",
            expected=(),
            actual=tuple(failures),
        )


def _body_field_signature(path: str | Path) -> tuple[int, int, int]:
    """Count fields in the main document part, excluding headers/footers."""
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    instructions = _body_field_instructions(path)
    return (
        sum(1 for node in root.iter(W + "fldChar") if node.get(W + "fldCharType") == "begin") + sum(1 for _ in root.iter(W + "fldSimple")),
        sum(1 for text in instructions if "ZOTERO_ITEM" in text),
        sum(1 for text in instructions if "ZOTERO_BIBL" in text),
    )


def _body_field_instructions(path: str | Path) -> tuple[str, ...]:
    """Reconstruct field instructions independently of run splitting."""
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    fields: list[list[str]] = []
    stack: list[list[str]] = []
    for element in root.iter():
        if element.tag == W + "fldSimple":
            fields.append([element.get(W + "instr", "")])
        elif element.tag == W + "fldChar":
            kind = element.get(W + "fldCharType")
            if kind == "begin":
                code: list[str] = []
                fields.append(code)
                stack.append(code)
            elif kind == "end" and stack:
                stack.pop()
        elif element.tag == W + "instrText":
            if stack:
                stack[-1].append(element.text or "")
            else:
                fields.append([element.text or ""])
    return tuple("".join(code).strip() for code in fields)


def _verify_saved_typography(path: str | Path | bytes, profile: JournalStyleProfile) -> None:
    """Resolve saved style inheritance before checking each formatted paragraph."""
    with zipfile.ZipFile(BytesIO(path) if isinstance(path, bytes) else path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
        styles = ET.fromstring(archive.read("word/styles.xml")) if "word/styles.xml" in archive.namelist() else ET.Element(W + "styles")
    expected_margins = {W + side: str(round(getattr(profile, "margin_" + side + "_inches") * 1440)) for side in ("top", "right", "bottom", "left")}
    sections = root.findall(f".//{W}sectPr")
    if not sections:
        raise VerificationFailed("saved document has no section margins to verify")
    for sect in sections:
        margins = sect.find(W + "pgMar")
        if margins is None or any(margins.get(key) != value for key, value in expected_margins.items()):
            raise VerificationFailed("saved document margins do not match the requested profile", expected=expected_margins, actual={} if margins is None else margins.attrib)

    def properties(element):
        return {} if element is None else {child.tag: dict(child.attrib) for child in element}

    def merge(*layers):
        result = {}
        for layer in layers:
            for tag, attrs in layer.items():
                result.setdefault(tag, {}).update(attrs)
        return result

    by_id = {node.get(W + "styleId"): node for node in styles.findall(W + "style")}
    default_style = next((key for key, node in by_id.items() if node.get(W + "type") == "paragraph" and node.get(W + "default") in ("1", "true", "on")), "Normal")
    defaults_p = properties(styles.find(f"{W}docDefaults/{W}pPrDefault/{W}pPr"))
    defaults_r = properties(styles.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr"))

    def style_properties(style_id, kind, seen=()):
        if style_id in seen:
            raise VerificationFailed("cyclic style inheritance in saved document")
        style = by_id.get(style_id)
        if style is None:
            return {}
        based = style.find(W + "basedOn")
        inherited = style_properties(based.get(W + "val"), kind, (*seen, style_id)) if based is not None else {}
        return merge(inherited, properties(style.find(W + kind)))

    paragraphs = root.findall(f".//{W}body/{W}p")
    table_paragraphs = root.findall(f".//{W}tbl//{W}p")
    first_body = next((p for p in paragraphs if any((n.text or "").strip() for n in p.iter(W + "t"))), None)
    for paragraph in [*paragraphs, *table_paragraphs]:
        text = "".join(node.text or "" for node in paragraph.iter(W + "t")).strip()
        if not text:
            continue
        in_table = paragraph in table_paragraphs
        ppr = paragraph.find(W + "pPr")
        style_node = None if ppr is None else ppr.find(W + "pStyle")
        style_id = default_style if style_node is None else style_node.get(W + "val", default_style)
        style = by_id.get(style_id)
        name_node = None if style is None else style.find(W + "name")
        style_name = (style_id if name_node is None else name_node.get(W + "val", style_id)).lower()
        role = classify_paragraph(text, style_name, paragraph is first_body, in_table)
        is_title, is_heading = role == "title", role == "heading"
        wanted_line = "240" if is_title or is_heading or profile.line_spacing_rule == 0 else "480"
        wanted_size = str(round((profile.title_size_pt if is_title else profile.heading_size_pt if is_heading else profile.font_size_pt) * 2))
        effective_p = merge(defaults_p, style_properties(style_id, "pPr"), properties(ppr))
        spacing = effective_p.get(W + "spacing", {})
        before = profile.title_space_before_pt if is_title else profile.heading_space_before_pt if is_heading else profile.space_before_pt
        after = profile.title_space_after_pt if is_title else profile.heading_space_after_pt if is_heading else profile.space_after_pt
        if (spacing.get(W + "line", "240") != wanted_line or spacing.get(W + "lineRule", "auto") != "auto"
            or spacing.get(W + "before", "0") != str(round(before * 20))
            or spacing.get(W + "after", "0") != str(round(after * 20))
            or spacing.get(W + "beforeAutospacing", "0") in ("1", "true", "on")
            or spacing.get(W + "afterAutospacing", "0") in ("1", "true", "on")):
            raise VerificationFailed("saved paragraph spacing does not match the requested profile", expected={"line": wanted_line, "before": before, "after": after}, actual={"text": text[:80], "spacing": spacing})
        if not in_table and effective_p.get(W + "jc", {}).get(W + "val", "left") not in ("left", "start"):
            raise VerificationFailed("saved paragraph alignment does not match the requested profile")
        inherited_r = merge(defaults_r, style_properties(style_id, "rPr"))
        for run in paragraph.iter(W + "r"):
            if not any(node.tag == W + "t" and (node.text or "") for node in run):
                continue
            rpr = run.find(W + "rPr")
            rstyle = None if rpr is None else rpr.find(W + "rStyle")
            effective_r = merge(inherited_r, style_properties(rstyle.get(W + "val"), "rPr") if rstyle is not None else {}, properties(rpr))
            font = effective_r.get(W + "rFonts", {}).get(W + "ascii")
            size = effective_r.get(W + "sz", {}).get(W + "val")
            if font != profile.font_name or size != wanted_size:
                raise VerificationFailed("saved run typography does not match the requested profile", expected={"font": profile.font_name, "size": wanted_size}, actual={"text": text[:80], "font": font, "size": size})


def _verify_protected_structure(
    before: Audit,
    after: Audit,
    *,
    allow_header_field_loss: bool = False,
) -> None:
    # Styling is allowed to change presentation XML (and, for anonymized
    # profiles, headers/footers), but it must not change document content or
    # any of the structures that Word/Zotero use to address that content.
    protected = (
        "body_text_sha256",
        "prose_text_sha256",
        "paragraph_count",
        "table_count",
        "tables",
        "bookmark_names",
        "bookmark_references",
        "comment_count",
        "live_comment_count",
        "orphan_comment_ids",
        "revision_count",
        "drawing_count",
        "media_count",
        "live_media_count",
        "dangling_media_targets",
    )
    for name in protected:
        if getattr(before, name) != getattr(after, name):
            raise VerificationFailed(
                f"protected invariant {name} changed during styling enforcement",
                expected=getattr(before, name),
                actual=getattr(after, name),
            )

    if allow_header_field_loss:
        before_body_fields = _body_field_signature(before.path)
        after_body_fields = _body_field_signature(after.path)
        if before_body_fields != after_body_fields or _body_field_instructions(before.path) != _body_field_instructions(after.path):
            raise VerificationFailed(
                "main-document field or citation content changed during header/footer stripping",
                expected=(before_body_fields, _body_field_instructions(before.path)),
                actual=(after_body_fields, _body_field_instructions(after.path)),
            )
    else:
        for name in ("field_count", "zotero_item_fields", "zotero_bibliography_fields"):
            if getattr(before, name) != getattr(after, name):
                raise VerificationFailed(
                    f"protected invariant {name} changed during styling enforcement",
                    expected=getattr(before, name),
                    actual=getattr(after, name),
                )

    # A package can retain the same object counts while becoming unreadable
    # (for example, a malformed complex field).  The audit is deliberately
    # the source of truth here; do not compare exact XML, which legitimately
    # changes when Word applies formatting or strips headers and footers.
    if after.structural_errors:
        raise VerificationFailed(
            "document has structural errors after styling enforcement",
            expected=(),
            actual=after.structural_errors,
        )


def _style_document_local(
    path: str | Path,
    profile_name: str = "radiology_ai",
    visible: bool = False,
    emphasis_policy: str = "preserve",
) -> dict[str, Any]:
    """Apply COM styling to one already-staged document and save it.

    This function is the operation worker entry point.  It deliberately does
    not stage, publish, or audit the source; the parent process owns those
    safety gates and passes only a disposable staging path here.
    """
    profile = PROFILES.get(profile_name)
    if not profile:
        raise RefusedUnsafe(
            f"unknown style profile '{profile_name}'. Available: {list(PROFILES.keys())}"
        )

    output_path = Path(path).resolve(strict=True)
    with IsolatedWordDocument(output_path, visible=visible) as session:
        doc = session.document

        # 1. Enforce Page Setup Margins (72 points = 1 inch)
        setup = doc.PageSetup
        setup.TopMargin = profile.margin_top_inches * 72.0
        setup.BottomMargin = profile.margin_bottom_inches * 72.0
        setup.LeftMargin = profile.margin_left_inches * 72.0
        setup.RightMargin = profile.margin_right_inches * 72.0

        # 2. Strip Headers & Footers if double-anonymized review requires no page numbers
        if profile.strip_header_footers:
            _strip_headers_footers(doc)

        # 3. Iterate Paragraphs to Apply Typography and Layout
        styled_count = 0
        first_body_seen = False
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            p_text = p.Range.Text.strip("\r\n\t \x07")
            if not p_text:
                continue

            # Table cell paragraphs receive the same body typography.  Their
            # existing run-level bold/italic properties and table geometry are
            # left untouched.
            in_table = False
            try:
                in_table = bool(p.Range.Information(12))  # wdWithInTable = 12
            except Exception:
                pass

            styled_count += 1
            rng = p.Range
            fmt = p.Format

            # Base font enforcement
            rng.Font.Name = profile.font_name

            # Alignment: Left-aligned
            if not in_table:
                fmt.Alignment = profile.alignment
            fmt.SpaceBeforeAuto = False
            fmt.SpaceAfterAuto = False

            # Heading vs Body Logic
            lower_text = p_text.lower()
            style_name = _style_name(p).lower()
            role = classify_paragraph(p_text, style_name, not first_body_seen, in_table)
            is_title, is_major_heading = role == "title", role == "heading"
            if not in_table:
                first_body_seen = True
            if is_title or is_major_heading:
                fmt.KeepWithNext = True

            if is_title:
                rng.Font.Size = profile.title_size_pt
                if emphasis_policy == "headings_and_summary":
                    rng.Font.Bold = True
                fmt.LineSpacingRule = 0  # Single spaced for title block
                fmt.SpaceBefore = profile.title_space_before_pt
                fmt.SpaceAfter = profile.title_space_after_pt
            elif is_major_heading:
                rng.Font.Size = profile.heading_size_pt
                if emphasis_policy == "headings_and_summary":
                    rng.Font.Bold = True
                fmt.LineSpacingRule = 0  # Single spaced for heading block
                fmt.SpaceBefore = profile.heading_space_before_pt
                fmt.SpaceAfter = profile.heading_space_after_pt
            elif p_text.startswith("Summary Statement:"):
                # Some journals require a bold "Summary Statement:" paragraph (journal-specific)
                rng.Font.Size = profile.font_size_pt
                if emphasis_policy == "headings_and_summary":
                    rng.Font.Bold = True
                _apply_line_spacing(fmt, profile)
                fmt.SpaceBefore = profile.space_before_pt
                fmt.SpaceAfter = profile.space_after_pt
            else:
                # Regular Body Paragraph
                rng.Font.Size = profile.font_size_pt
                _apply_line_spacing(fmt, profile)
                fmt.SpaceBefore = profile.space_before_pt
                fmt.SpaceAfter = profile.space_after_pt

        session.save()

    return {
        "output": str(output_path),
        "profile": profile_name,
        "paragraphs_styled": styled_count,
    }


def apply_journal_style(
    source: str | Path,
    output: str | Path | None = None,
    profile_name: str = "radiology_ai",
    *,
    visible: bool = False,
    emphasis_policy: str = "preserve",
    source_sha256: str | None = None,
) -> dict[str, Any]:
    """Apply journal typography through a bounded isolated Word operation."""
    profile = PROFILES.get(profile_name)
    if not profile:
        raise RefusedUnsafe(
            f"unknown style profile '{profile_name}'. Available: {list(PROFILES.keys())}"
        )

    if emphasis_policy not in ("preserve", "headings_and_summary"):
        raise RefusedUnsafe("Unknown emphasis_policy")
    resolved_source = Path(source).resolve()
    target_output = Path(output).resolve() if output is not None else resolved_source
    before = audit_docx(resolved_source)
    if source_sha256 is not None and source_sha256.upper() != before.sha256:
        raise RefusedUnsafe("Source changed since inspection; inspect again")
    before_instructions = _body_field_instructions(resolved_source)
    from .style_audit import snapshot_emphasis
    before_emphasis = snapshot_emphasis(resolved_source)

    # Imported lazily so the style module remains usable in environments that
    # only perform offline audits, while the worker owns the COM lifecycle.
    from .operation_process import run_operation

    with staged_document(resolved_source, target_output) as output_path:
        if audit_docx(output_path).sha256 != before.sha256:
            raise RefusedUnsafe("Source changed while staging; inspect again")
        worker_result = run_operation(
            "style",
            {
                "source": str(output_path),
                "profile_name": profile_name,
                "visible": visible,
                "emphasis_policy": emphasis_policy,
            },
            timeout_seconds=120,
        )

        after = audit_docx(output_path)
        if _body_field_instructions(output_path) != before_instructions:
            raise VerificationFailed("styling changed main-document field instructions")
        after_emphasis = snapshot_emphasis(output_path)
        if emphasis_policy == "preserve" and after_emphasis != before_emphasis:
            raise VerificationFailed("layout styling changed bold or italic outside its declared scope")
        if emphasis_policy == "headings_and_summary":
            from .style_audit import audit_document_style
            role_audit = audit_document_style(resolved_source, _unbounded=True)
            allowed_paragraphs = {p["index"] for p in role_audit["paragraphs"] if p["role"] in ("title", "heading", "summary")}
            if before_emphasis.keys() != after_emphasis.keys():
                raise VerificationFailed("semantic styling changed visible character layout")
            for key, (text, bold, italic) in before_emphasis.items():
                expected = (text, True if key[0] in allowed_paragraphs else bold, italic)
                if after_emphasis[key] != expected:
                    raise VerificationFailed("semantic styling changed emphasis outside the declared paragraph roles", expected=expected, actual=after_emphasis[key])
        _verify_saved_typography(output_path, profile)
        _verify_protected_structure(
            before,
            after,
            allow_header_field_loss=profile.strip_header_footers,
        )

    return {
        "source": str(resolved_source),
        "source_sha256": before.sha256,
        "output": str(target_output),
        "output_sha256": after.sha256,
        "profile": profile_name,
        "paragraphs_styled": worker_result.get("paragraphs_styled", 0),
        "margins": {
            "top_inches": profile.margin_top_inches,
            "bottom_inches": profile.margin_bottom_inches,
            "left_inches": profile.margin_left_inches,
            "right_inches": profile.margin_right_inches,
        },
        "typography": {
            "font": profile.font_name,
            "font_size_pt": profile.font_size_pt,
            "double_spaced": (profile.line_spacing_rule == 2),
            "single_spaced": (profile.line_spacing_rule == 0),
            "space_before_pt": profile.space_before_pt,
            "space_after_pt": profile.space_after_pt,
        },
        "protected_structure_unchanged": True,
        "emphasis_policy": emphasis_policy,
    }
