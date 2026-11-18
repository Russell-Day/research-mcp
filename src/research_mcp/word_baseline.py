"""Word's own save of an untouched document as the reference for verifying a Word pass (plan 973 S1).

Word rewrites some XML on every save (ids, run splits) and, when the body ends with a table, appends an empty
paragraph after it: a body may not end with a table in Word's own form. A tool that checks its saved result
against the source would then report "paragraph count changed 117 -> 118" for an edit that changed nothing there.
``edit_document`` compares against Word's save of the untouched document instead (``editing._replan_on_baseline``);
``format_tables`` and ``apply_document_style`` do the same through ``choose_reference``.

The baseline is adopted only after ``sections.baseline_preservation_errors`` shows it preserves the source (same
text, every character's formatting, paragraph/table/row/cell properties, bookmark and comment extents, fields,
content controls, hyperlinks). The one difference tolerated beyond that check is the empty paragraph Word appends
after a body-final table: it is trimmed from a temporary copy of the baseline before the check, and only when it
holds no run, field, bookmark or comment marker and its properties hold no section break, numbering, frame or
revision. Section properties are compared too (``section_errors``): every ``w:sectPr``, body-final and in
paragraphs, with the header and footer parts it references. Otherwise the caller compares against the source. When Word's save
of the untouched document is not faithful for another reason (it drops properties it does not read, such as an
``rFonts`` written after ``sz``) but did append that paragraph, the source is compared in Word's body-final form:
a copy of the source with that empty paragraph, taken verbatim from the baseline, after its final table.
"""

from __future__ import annotations

import copy
import hashlib
from dataclasses import dataclass
from pathlib import Path
import posixpath
import re
import tempfile
import xml.etree.ElementTree as ET
import zipfile

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{" + W_NS + "}"
DOCUMENT_PART = "word/document.xml"
WORD_BASELINE = "word_baseline"
SOURCE = "source"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
R_ID = R_NS + "id"
RELATIONSHIPS_PART = "word/_rels/document.xml.rels"
# What the paragraph Word appends after a body-final table may hold: its properties and nothing else, and in its
# properties nothing that makes it more than an empty line (a section break, list numbering, a frame, a revision).
_APPENDED_PARAGRAPH_CHILDREN = {W + "pPr"}
_STRUCTURAL_PROPERTIES = {W + name for name in (
    "sectPr", "numPr", "framePr", "pPrChange", "rPrChange", "ins", "del", "moveFrom", "moveTo")}


@dataclass(frozen=True)
class Reference:
    """What a Word pass is verified against: ``path`` and ``compared_against`` ("word_baseline" or "source");
    ``problems`` says why the baseline was not adopted (empty when it was)."""

    path: Path
    compared_against: str
    problems: tuple[str, ...] = ()


def _body_children(path: str | Path) -> list[ET.Element]:
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read(DOCUMENT_PART))
    body = root.find(W + "body")
    return [] if body is None else [child for child in body if child.tag != W + "sectPr"]


def _is_empty_paragraph(node: ET.Element) -> bool:
    return (node.tag == W + "p" and all(child.tag in _APPENDED_PARAGRAPH_CHILDREN for child in node)
            and not any(element.tag in _STRUCTURAL_PROPERTIES for element in node.iter()))


def appended_final_paragraph(source: str | Path, baseline: str | Path) -> bool:
    """True when the source body ends with a table and the baseline's body is the same plus one empty paragraph
    after it, as Word writes a body-final table."""

    old, new = _body_children(source), _body_children(baseline)
    return (bool(old) and old[-1].tag == W + "tbl" and len(new) == len(old) + 1
            and new[-2].tag == W + "tbl" and _is_empty_paragraph(new[-1]))


_PARAGRAPH_START = re.compile(r"<w:p(?=[\s/>])")


def _tail(xml: str) -> int:
    """Where the body's content ends in ``xml`` (document.xml as text): its final sectPr, else ``</w:body>``."""

    body_end = xml.rindex("</w:body>")
    sect = xml.rfind("<w:sectPr", 0, body_end)
    return sect if sect != -1 and xml.rfind("</w:p>", 0, body_end) < sect else body_end


def _final_paragraph_span(xml: str) -> tuple[int, int]:
    """Start and end of the body's last paragraph, which ``appended_final_paragraph`` has shown is empty: from its
    start tag to the body's sectPr (or ``</w:body>``)."""

    tail = _tail(xml)
    starts = [match.start() for match in _PARAGRAPH_START.finditer(xml, 0, tail)]
    if not starts:
        raise ValueError("no paragraph before the end of the body")
    return starts[-1], tail


def _without_final_paragraph(xml: str) -> str:
    """``xml`` (document.xml as text) without the body's last paragraph; nothing else is cut."""

    start, end = _final_paragraph_span(xml)
    return xml[:start] + xml[end:]


def trimmed_copy(baseline: str | Path, target: str | Path) -> Path:
    """A copy of ``baseline`` without its appended final paragraph; every other part byte-identical."""

    target = Path(target)
    with zipfile.ZipFile(baseline) as source, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == DOCUMENT_PART:
                data = _without_final_paragraph(data.decode("utf-8")).encode("utf-8")
            out.writestr(info, data)
    old, new = _body_children(baseline), _body_children(target)
    if len(new) != len(old) - 1 or [n.tag for n in new] != [n.tag for n in old[:-1]]:
        raise ValueError("trimming the appended paragraph changed other body elements")
    return target


_PREFIXED_ATTRIBUTE = re.compile(r'\s+([A-Za-z][\w.-]*):[\w.-]+="[^"]*"')


def _declared_only(paragraph: str, xml: str) -> str:
    """``paragraph`` with the attributes whose namespace prefix ``xml``'s root does not declare left out (Word adds
    ``w14:paraId``/``w14:textId`` to a paragraph it writes); a prefixed element the root does not declare is an
    error."""

    root_tag = xml[xml.index("<w:document"):xml.index(">", xml.index("<w:document"))]
    declared = set(re.findall(r"xmlns:([\w.-]+)=", root_tag)) | {"xml"}
    kept = _PREFIXED_ATTRIBUTE.sub(lambda m: m.group(0) if m.group(1) in declared else "", paragraph)
    undeclared = {prefix for prefix in re.findall(r"</?([\w.-]+):", kept) if prefix not in declared}
    if undeclared:
        raise ValueError(f"Word's appended paragraph uses namespaces the source does not declare: {sorted(undeclared)}")
    return kept


def source_in_word_form(source: str | Path, baseline: str | Path, target: str | Path) -> Path:
    """A copy of ``source`` with the empty paragraph Word appended after its final table in ``baseline`` (verbatim)
    inserted at the same place; every other part and every other byte of document.xml unchanged. Only for a
    baseline where ``appended_final_paragraph(source, baseline)`` holds."""

    if not appended_final_paragraph(source, baseline):
        raise ValueError("the baseline does not end with Word's appended paragraph")
    with zipfile.ZipFile(baseline) as archive:
        xml = archive.read(DOCUMENT_PART).decode("utf-8")
    start, end = _final_paragraph_span(xml)
    paragraph = xml[start:end]
    target = Path(target)
    with zipfile.ZipFile(source) as old, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for info in old.infolist():
            data = old.read(info.filename)
            if info.filename == DOCUMENT_PART:
                xml = data.decode("utf-8")
                at = _tail(xml)
                data = (xml[:at] + _declared_only(paragraph, xml) + xml[at:]).encode("utf-8")
            out.writestr(info, data)
    if not appended_final_paragraph(source, target):
        raise ValueError("adding Word's appended paragraph changed other body elements")
    return target


_EXTENT = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent"


def _story(content: ET.Element, linked: dict[str, str]) -> tuple:
    """What a header or footer part shows, read as the body checks read paragraphs (``word_tracked._paragraph``:
    canonical paragraph properties, then each character with its run properties as written, rsids and revision
    records aside), plus per paragraph its field instructions, the targets its relationship ids name (``linked``:
    the hash of an image's bytes, an external link's target) and each drawing's extent; per table its properties
    and grid. Ids, rsids, proofing marks and the drawing defaults Word writes on save (``wp14``/``a14`` attributes,
    docPr ids, ``w:noProof`` on a picture's run) are not part of it."""

    from .word_tracked import _paragraph, canonical as properties

    paragraphs = []
    for paragraph in content.iter(W + "p"):
        instructions = "".join(node.text or "" for node in paragraph.iter(W + "instrText"))
        fields = [" ".join(instructions.split())] if instructions.strip() else []
        fields += [" ".join((node.get(W + "instr") or "").split()) for node in paragraph.iter(W + "fldSimple")]
        # Document order of the linking elements counts (two pictures swapped is a change);
        # only the r: attributes within one element are sorted, since Word may reorder them.
        links = [tuple(sorted(linked.get(value, f"unknown relationship {value}")
                              for key, value in node.attrib.items() if key.startswith(R_NS)))
                 for node in paragraph.iter() if any(key.startswith(R_NS) for key in node.attrib)]
        extents = [(node.get("cx"), node.get("cy")) for node in paragraph.iter(_EXTENT)]
        paragraphs.append((_paragraph(paragraph), fields, links, extents))
    tables = [(properties(table.find(W + "tblPr")), properties(table.find(W + "tblGrid")))
              for table in content.iter(W + "tbl")]
    return tuple(paragraphs), tuple(tables)


def _sections(path: str | Path) -> list[str]:
    """Every ``w:sectPr`` of the body in document order (in paragraphs, then the body-final one), canonical (ids and
    rsids aside), with each header/footer reference's relationship id replaced by what the part it names shows
    (``_story``; each relationship id inside that part by the hash of the media it names or its external target):
    Word may renumber relationships, rename parts and rewrite drawing markup on save, never change what they hold."""

    from .sections import canonical

    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read(DOCUMENT_PART))
        names = set(archive.namelist())
        targets: dict[str, str] = {}
        if RELATIONSHIPS_PART in names:
            for relationship in ET.fromstring(archive.read(RELATIONSHIPS_PART)):
                targets[relationship.get("Id", "")] = relationship.get("Target", "")

        def resolve(folder: str, target: str) -> str:
            return target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))

        def part(relationship_id: str | None) -> str:
            target = targets.get(relationship_id or "")
            if not target:
                return f"missing relationship {relationship_id}"
            name = resolve("word", target)
            if name not in names:
                return f"missing part {name}"
            # What the part's own relationships point at (a header image, a hyperlink), by content: a relationship
            # id in the part is replaced by the hash of the bytes it names, or by an external target as written.
            folder, base = posixpath.split(name)
            rels_name = posixpath.join(folder, "_rels", base + ".rels")
            linked: dict[str, str] = {}
            if rels_name in names:
                for relationship in ET.fromstring(archive.read(rels_name)):
                    link = relationship.get("Target", "")
                    if relationship.get("TargetMode") == "External":
                        linked[relationship.get("Id", "")] = "external " + link
                    else:
                        media = resolve(folder, link)
                        linked[relationship.get("Id", "")] = ("sha256 " + hashlib.sha256(archive.read(media)).hexdigest()
                                                             if media in names else f"missing part {media}")
            return repr(_story(ET.fromstring(archive.read(name)), linked))

        body = root.find(W + "body")
        found = []
        for section in ([] if body is None else body.iter(W + "sectPr")):
            clone = copy.deepcopy(section)
            for reference in clone:
                if reference.tag in (W + "headerReference", W + "footerReference"):
                    reference.set("part", part(reference.attrib.pop(R_ID, None)))
            found.append(canonical(clone))
    return found


def section_errors(source: str | Path, baseline: str | Path) -> list[str]:
    """Differences in section properties (page size, margins, columns, headers and footers...) between ``source``
    and ``baseline``; ``sections.baseline_preservation_errors`` does not read the body-final ``w:sectPr``."""

    old, new = _sections(source), _sections(baseline)
    if len(old) != len(new):
        return [f"section count changed ({len(old)} -> {len(new)})"]
    return [f"section {number} properties, headers or footers changed"
            for number, (before, after) in enumerate(zip(old, new), start=1) if before != after]


def preservation_errors(source: str | Path, baseline: str | Path) -> list[str]:
    """``sections.baseline_preservation_errors`` plus ``section_errors``, tolerating the empty paragraph Word
    appends after a body-final table."""

    from .sections import baseline_preservation_errors

    if not appended_final_paragraph(source, baseline):
        return [*baseline_preservation_errors(source, baseline), *section_errors(source, baseline)]
    with tempfile.TemporaryDirectory(prefix="research-mcp-baseline-trim-") as folder:
        try:
            trimmed = trimmed_copy(baseline, Path(folder) / "trimmed.docx")
        except (ValueError, KeyError, UnicodeDecodeError) as exc:
            return [f"Word's appended final paragraph could not be set aside: {exc}"]
        return [*baseline_preservation_errors(source, trimmed), *section_errors(source, trimmed)]


def choose_reference(source: str | Path, baseline: str | Path | None) -> Reference:
    """Word's save of the untouched ``source`` when it preserves it, else the source: in Word's body-final form
    (``source_in_word_form``, written beside the baseline) when the baseline appended the empty paragraph after a
    final table, else the source file itself."""

    source = Path(source)
    if baseline is None or not Path(baseline).exists():
        return Reference(source, SOURCE, ("no Word baseline was saved",))
    try:
        problems = preservation_errors(source, baseline)
    except Exception as exc:  # an unreadable baseline is not adopted; the source check still runs
        problems = [f"the baseline could not be read: {type(exc).__name__}: {exc}"]
    if not problems:
        return Reference(Path(baseline), WORD_BASELINE)
    try:
        if appended_final_paragraph(source, baseline):
            word_form = source_in_word_form(source, baseline, Path(baseline).with_name("source_word_form.docx"))
            return Reference(word_form, SOURCE, tuple(problems))
    except Exception as exc:
        problems = [*problems, f"Word's appended paragraph could not be added to the source: {exc}"]
    return Reference(source, SOURCE, tuple(problems))
