"""
document_processor.py
Core find-and-replace logic for Word documents (.docx).

Design note — one traversal, one source of truth
------------------------------------------------
Paragraph text, character offsets, and the XML nodes that hold the text are all
produced by a single walk of the document XML (`_paragraph_text_and_nodes`).
They are therefore structurally aligned and cannot drift apart.

This matters because python-docx's convenience APIs disagree with each other:
`Paragraph.text` is built from ``w:r | w:hyperlink`` while `Paragraph.runs`
returns only direct ``w:r`` children. Searching one and editing the other puts
replacements at the wrong offset in any paragraph containing a hyperlink.

Paragraphs are identified by a structural key (part name + ordinal position in
that part's walk) rather than by a positional counter over python-docx objects.
The key is derived from the XML tree, so it is stable across the separate
`Document()` loads used by the scan and apply phases, and a physical paragraph
is visited exactly once even when python-docx would surface it several times
(merged table cells, headers shared by several sections).

Replacements edit ``w:t`` element text in place. Assigning to `Run.text` would
call ``clear_content()``, which strips every child of the run — destroying
inline images, footnote references, comment anchors and field characters that
happen to share a run with matched text.

Covered content: body paragraphs, tables (including nested tables), text boxes,
headers and footers, footnotes, endnotes and comments. Text inside hyperlinks,
tracked insertions (``w:ins``), smart tags and content controls is included. See
README "Known Limitations" for what is deliberately not covered.

Offsets recorded by a scan describe one version of a file. `apply_changes`
re-checks each file's size and mtime before touching it and refuses a file that
changed in between, because a stale offset can still validate against different
text.
"""

import copy
import csv
import getpass
import os
import re
import shutil
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Dict, List, NamedTuple, Optional, Sequence, Tuple, Union

from docx import Document
from docx.opc.constants import CONTENT_TYPE as _CT
from docx.opc.exceptions import PackageNotFoundError
from docx.opc.part import PartFactory, XmlPart
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Footnote, endnote and comment parts must load as XML so their paragraphs can
# be walked and edited; otherwise they arrive as opaque blobs and are silently
# skipped. python-docx registers a class for comments from 1.2.0 and never for
# footnotes or endnotes, so all three are registered here rather than relying on
# which version pip resolved - the supported range starts at 1.1.0, which
# registers none of them.
#
# `PartFactory.part_type_for` is the documented extension point; setdefault so a
# python-docx that ships its own class for any of these keeps precedence.
PartFactory.part_type_for.setdefault(_CT.WML_FOOTNOTES, XmlPart)
PartFactory.part_type_for.setdefault(_CT.WML_ENDNOTES, XmlPart)
PartFactory.part_type_for.setdefault(_CT.WML_COMMENTS, XmlPart)


# --------------------------------------------------------------------------
# XML namespace helpers
# --------------------------------------------------------------------------

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_XML = "http://www.w3.org/XML/1998/namespace"


def _w(tag: str) -> str:
    return "{%s}%s" % (NS_W, tag)


def _mc(tag: str) -> str:
    return "{%s}%s" % (NS_MC, tag)


_XML_SPACE = "{%s}space" % NS_XML

_REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Parts that hold their own paragraphs, reached by relationship from the
# document part: (relationship type, wrapper element, location type, label).
_NOTE_PARTS = (
    (f"{_REL_BASE}/footnotes", "footnote", 'footnote', 'Footnote'),
    (f"{_REL_BASE}/endnotes", "endnote", 'endnote', 'Endnote'),
    (f"{_REL_BASE}/comments", "comment", 'comment', 'Comment'),
)

# Separator "footnotes" are Word's rule lines, not content.
_NOTE_SKIP_TYPES = frozenset({'separator', 'continuationSeparator'})

_T = _w('t')
_P = _w('p')
_TBL = _w('tbl')
_TR = _w('tr')
_TC = _w('tc')
_SDT = _w('sdt')
_SDT_CONTENT = _w('sdtContent')
_TXBX_CONTENT = _w('txbxContent')
_ALT_CONTENT = _mc('AlternateContent')
_CHOICE = _mc('Choice')
_FALLBACK = _mc('Fallback')

# Leaf elements that contribute characters to a paragraph's text.
# w:t text is read from the element; the others render as a fixed character.
_ATOMIC_LEAVES = {
    _w('tab'): '\t',
    _w('ptab'): '\t',
    _w('br'): '\n',
    _w('cr'): '\n',
    _w('noBreakHyphen'): '-',
}

# Inline containers descended into when gathering a paragraph's text.
# This is an allowlist: anything not named here is deliberately NOT entered.
# Notably excluded are w:drawing / w:pict / w:object (a text box living inside a
# run is collected separately, as its own paragraphs), mc:AlternateContent,
# w:del and w:moveFrom (deleted text, which must not be edited), and
# w:fldSimple (a cached field result that Word regenerates).
_INLINE_CONTAINERS = frozenset({
    _w('r'), _w('hyperlink'), _w('ins'), _w('moveTo'),
    _w('smartTag'), _w('sdt'), _w('sdtContent'),
    _w('bdo'), _w('dir'), _w('customXml'),
})


@dataclass(frozen=True)
class Rule:
    """One find/replace pair. A scan runs a list of these as a set."""
    find_text: str
    replace_text: str = ""


@dataclass
class Match:
    """Represents a single find-text occurrence in a document."""
    file_path: str
    location_type: str          # 'body', 'table', 'header', 'footer',
                                # 'footnote', 'endnote' or 'comment'
    location_detail: str        # e.g., 'Paragraph 5', 'Table 2, Row 1, Cell 3'
    paragraph_key: str          # Structural key, stable across document loads
    char_offset: int            # Character offset within the paragraph's full text
    context_before: str         # ~50 chars before match
    match_text: str             # The matched text
    context_after: str          # ~50 chars after match
    is_selected: bool = True    # User's checkbox state
    applied: Optional[bool] = None   # Set by apply_changes(): True, False, or None
    file_mtime: float = 0.0     # File mtime when scanned, for the staleness check
    file_size: int = 0          # File size when scanned, for the staleness check
    find_text: str = ""         # The rule's search text ("" -> use match_text)
    replace_text: str = ""      # The rule's replacement text
    section: str = ""           # Nearest heading path, e.g. "GENERAL > REFERENCES"
    conflict: bool = False      # Overlaps a match from another rule; never applied
    skip_reason: str = ""       # Why apply_changes() did not apply it, if it did not

    @property
    def display_context(self) -> str:
        """Format the context for display in the GUI (single line)."""
        before = self.context_before.replace("\n", "⏎ ").replace("\t", " ")
        after = self.context_after.replace("\n", " ⏎").replace("\t", " ")
        match = self.match_text.replace("\n", "⏎").replace("\t", " ")
        return f"{before}[{match}]{after}"

    @property
    def effective_find_text(self) -> str:
        """The text this match was found with (hand-built matches carry none)."""
        return self.find_text or self.match_text


@dataclass
class FileResult:
    """All matches for a single file."""
    file_path: str
    file_name: str              # Relative to the scanned folder
    matches: List[Match] = field(default_factory=list)
    error: Optional[str] = None

    @property
    def match_count(self) -> int:
        return len(self.matches)

    @property
    def selected_count(self) -> int:
        return sum(1 for m in self.matches if m.is_selected)

    @property
    def conflict_count(self) -> int:
        return sum(1 for m in self.matches if m.conflict)

    @property
    def selectable_count(self) -> int:
        """Matches the user is allowed to select: everything but conflicts."""
        return sum(1 for m in self.matches if not m.conflict)


@dataclass
class _TextNode:
    """One text-bearing leaf element and the slice of paragraph text it owns."""
    element: object
    start: int
    length: int
    editable: bool      # True for w:t; False for atomic leaves (w:tab, w:br, ...)


# --------------------------------------------------------------------------
# Single-pass paragraph text extraction
# --------------------------------------------------------------------------

def _paragraph_text_and_nodes(p_el) -> Tuple[str, List[_TextNode]]:
    """
    Walk a ``w:p`` element and return (full_text, nodes).

    ``nodes`` maps character ranges of ``full_text`` onto the XML elements that
    produced them, built in the same pass as the text itself.
    """
    parts: List[str] = []
    nodes: List[_TextNode] = []
    pos = 0

    def visit(el):
        nonlocal pos
        for child in el:
            tag = child.tag
            if tag == _T:
                text = child.text or ""
                if text:
                    nodes.append(_TextNode(child, pos, len(text), True))
                    parts.append(text)
                    pos += len(text)
            elif tag in _ATOMIC_LEAVES:
                ch = _ATOMIC_LEAVES[tag]
                nodes.append(_TextNode(child, pos, len(ch), False))
                parts.append(ch)
                pos += len(ch)
            elif tag in _INLINE_CONTAINERS:
                visit(child)

    visit(p_el)
    return "".join(parts), nodes


def paragraph_text(p_el) -> str:
    """Public helper: the text of a ``w:p`` element as this module sees it."""
    return _paragraph_text_and_nodes(p_el)[0]


# --------------------------------------------------------------------------
# Structural traversal: locating every paragraph exactly once
# --------------------------------------------------------------------------

def _has_intervening(el, ancestor_tags, stop_el) -> bool:
    """True if any ancestor of `el` below `stop_el` has one of `ancestor_tags`."""
    parent = el.getparent()
    while parent is not None and parent is not stop_el:
        if parent.tag in ancestor_tags:
            return True
        parent = parent.getparent()
    return False


# --------------------------------------------------------------------------
# Section context: the nearest preceding headings of a body paragraph
# --------------------------------------------------------------------------

_HEADING_LABEL_MAX = 60


def _styles_root(doc):
    """
    The ``w:styles`` element, or None if the document has no styles part.

    Resolved through the document part's relationships, never through
    `doc.styles`: that property *creates* a default styles part when one is
    missing, and a scan must not change what gets saved.
    """
    for rel in doc.part.rels.values():
        if rel.reltype == f"{_REL_BASE}/styles" and not rel.is_external:
            return getattr(rel.target_part, 'element', None)
    return None


def _outline_levels_by_style(styles_root) -> Dict[str, Optional[int]]:
    """
    Map every paragraph style id to its outline level (0-8), or None.

    A style's own ``w:outlineLvl`` wins; level 9 means "body text" and stops
    inheritance - Word's own TOC Heading style is based on Heading 1 with an
    explicit level 9. Otherwise the ``w:basedOn`` chain is followed. A style
    with no level anywhere but named "heading N" is treated as level N-1, for
    templates whose heading styles were stripped of outline levels.
    """
    info: Dict[str, Tuple[str, Optional[str], Optional[int]]] = {}
    if styles_root is None:
        return {}
    for style in styles_root.findall(_w('style')):
        if style.get(_w('type')) != 'paragraph':
            continue
        style_id = style.get(_w('styleId'))
        if not style_id:
            continue
        name_el = style.find(_w('name'))
        name = (name_el.get(_w('val')) if name_el is not None else "") or ""
        based_el = style.find(_w('basedOn'))
        based_on = based_el.get(_w('val')) if based_el is not None else None
        own = None
        lvl = style.find(f"{_w('pPr')}/{_w('outlineLvl')}")
        if lvl is not None:
            try:
                own = int(lvl.get(_w('val')))
            except (TypeError, ValueError):
                own = None
        info[style_id] = (name, based_on, own)

    resolved: Dict[str, Optional[int]] = {}

    def resolve(style_id: str, seen: set) -> Optional[int]:
        if style_id in resolved:
            return resolved[style_id]
        if style_id not in info or style_id in seen:
            return None
        seen.add(style_id)
        name, based_on, own = info[style_id]
        if own is not None:
            level = own if 0 <= own <= 8 else None
        else:
            level = resolve(based_on, seen) if based_on else None
            if level is None:
                m = re.fullmatch(r'heading ([1-9])', name.strip(), re.IGNORECASE)
                if m:
                    level = int(m.group(1)) - 1
        resolved[style_id] = level
        return level

    for style_id in info:
        resolve(style_id, set())
    return resolved


def _outline_level(p_el, levels: Dict[str, Optional[int]]) -> Optional[int]:
    """A paragraph's outline level: its own ``w:outlineLvl``, else its style's."""
    ppr = p_el.find(_w('pPr'))
    if ppr is None:
        return None
    lvl = ppr.find(_w('outlineLvl'))
    if lvl is not None:
        try:
            n = int(lvl.get(_w('val')))
        except (TypeError, ValueError):
            n = None
        if n is not None:
            return n if 0 <= n <= 8 else None
    pstyle = ppr.find(_w('pStyle'))
    if pstyle is not None:
        return levels.get(pstyle.get(_w('val')))
    return None


class _HeadingTracker:
    """
    Follows headings through a body walk so each paragraph can record the
    headings above it, e.g. ``GENERAL > REFERENCES``.

    Only flow-level body paragraphs are observed; paragraphs in table cells
    and text boxes inherit the current path without changing it. Numbering
    that Word generates ("PART 1", "1.02") is not part of the paragraph text
    and so does not appear in the label.
    """

    def __init__(self, levels: Dict[str, Optional[int]]):
        self.levels = levels
        self.stack: List[Tuple[int, str]] = []

    def observe(self, p_el) -> None:
        level = _outline_level(p_el, self.levels)
        if level is None:
            return
        label = " ".join(paragraph_text(p_el).split())
        if not label:
            return
        if len(label) > _HEADING_LABEL_MAX:
            label = label[:_HEADING_LABEL_MAX - 1].rstrip() + "…"
        self.stack = [(lvl, text) for lvl, text in self.stack if lvl < level]
        self.stack.append((level, label))

    @property
    def path(self) -> str:
        return " > ".join(text for _lvl, text in self.stack)


def _walk_textboxes(p_el, loc_type: str, para_detail: str, out: List[tuple],
                    headings: Optional[_HeadingTracker] = None) -> None:
    """Collect paragraphs of any text box anchored in this paragraph."""
    for txbx in p_el.iter(_TXBX_CONTENT):
        # Skip the legacy mc:Fallback copy of a text box; Word renders mc:Choice.
        if _has_intervening(txbx, {_FALLBACK}, p_el):
            continue
        # A nested text box is reached through its own paragraph's walk.
        if _has_intervening(txbx, {_TXBX_CONTENT}, p_el):
            continue
        _walk_container(txbx, loc_type, f"{para_detail}, Text Box", out, headings)


def _walk_table(tbl_el, loc_type: str, prefix: str, table_no: int, out: List[tuple],
                headings: Optional[_HeadingTracker] = None) -> None:
    """Walk a ``w:tbl``. Iterating XML visits each physical cell exactly once."""
    base = f"{prefix}, " if prefix else ""
    cell_type = 'table' if loc_type == 'body' else loc_type
    row_no = 0
    for row in tbl_el:
        if row.tag != _TR:
            continue
        row_no += 1
        cell_no = 0
        for cell in row:
            if cell.tag != _TC:
                continue
            cell_no += 1
            cell_prefix = f"{base}Table {table_no}, Row {row_no}, Cell {cell_no}"
            _walk_container(cell, cell_type, cell_prefix, out, headings)


def _walk_container(el, loc_type: str, prefix: str, out: List[tuple],
                    headings: Optional[_HeadingTracker] = None) -> None:
    """
    Walk a block-level container (``w:body``, ``w:hdr``, ``w:ftr``, ``w:tc``,
    ``w:txbxContent``) in document order, appending
    ``(p_element, location_type, location_detail, section)`` tuples to `out`.

    `headings` tracks the heading path for the body part; flow-level body
    paragraphs (an empty prefix) update it, everything else only reads it.
    """
    multi_para = sum(1 for c in el if c.tag == _P) > 1
    para_no = 0
    table_no = 0

    for child in el:
        tag = child.tag
        if tag == _P:
            para_no += 1
            if not prefix:
                detail = f"Paragraph {para_no}"
            elif multi_para:
                detail = f"{prefix}, Para {para_no}"
            else:
                detail = prefix
            if headings is not None and loc_type == 'body' and not prefix:
                headings.observe(child)
            section = headings.path if headings is not None else ""
            out.append((child, loc_type, detail, section))
            _walk_textboxes(child, loc_type, detail, out, headings)
        elif tag == _TBL:
            table_no += 1
            _walk_table(child, loc_type, prefix, table_no, out, headings)
        elif tag == _SDT:
            content = child.find(_SDT_CONTENT)
            if content is not None:
                _walk_container(content, loc_type, prefix, out, headings)
        elif tag == _ALT_CONTENT:
            choice = child.find(_CHOICE)
            if choice is not None:
                _walk_container(choice, loc_type, prefix, out, headings)


class _Located(NamedTuple):
    """One paragraph found by the traversal."""
    element: object
    loc_type: str
    detail: str
    key: str
    section: str


def _locate_paragraphs(doc) -> List[_Located]:
    """
    Collect every paragraph in the document exactly once.

    `key` is ``"<part name>#<ordinal>"``. The ordinal is the position in this
    deterministic XML walk of that part, so the same physical paragraph
    receives the same key on every load of an unchanged file. `section` is
    the heading path above a body paragraph, "" elsewhere.
    """
    collected: List[_Located] = []

    def add_part(part_name: str, root_el, loc_type: str, label: str,
                 headings: Optional[_HeadingTracker] = None) -> None:
        raw: List[tuple] = []
        _walk_container(root_el, loc_type, label, raw, headings)
        for ordinal, (p_el, l_type, detail, section) in enumerate(raw):
            collected.append(_Located(p_el, l_type, detail,
                                      f"{part_name}#{ordinal}", section))

    seen_parts = set()

    # --- Body (includes tables, nested tables and text boxes)
    try:
        levels = _outline_levels_by_style(_styles_root(doc))
    except Exception:
        levels = {}
    add_part(str(doc.part.partname), doc.element.body, 'body', "",
             _HeadingTracker(levels))

    # --- Headers and footers.
    # Resolved through w:headerReference / w:footerReference rather than
    # section.header, which creates an empty header part as a side effect.
    # A section linked to the previous one carries no reference of its own, so
    # a shared header part is visited once.
    ref_kinds = (
        (_w('headerReference'), 'header', 'Header'),
        (_w('footerReference'), 'footer', 'Footer'),
    )
    type_labels = {'default': '', 'first': 'First Page ', 'even': 'Even Page '}

    for section in doc.sections:
        try:
            sect_pr = section._sectPr
        except Exception:
            continue
        for ref_tag, loc_type, base_label in ref_kinds:
            for ref in sect_pr.findall(ref_tag):
                r_id = ref.get(qn('r:id'))
                if not r_id:
                    continue
                try:
                    part = doc.part.related_parts[r_id]
                except Exception:
                    continue
                part_name = str(part.partname)
                if part_name in seen_parts:
                    continue
                seen_parts.add(part_name)
                hf_type = ref.get(_w('type')) or 'default'
                label = f"{type_labels.get(hf_type, '')}{base_label}"
                try:
                    root_el = part.element
                except Exception:
                    continue
                add_part(part_name, root_el, loc_type, label)

    # --- Footnotes, endnotes and comments, each a part of its own.
    for reltype, wrapper, loc_type, base_label in _NOTE_PARTS:
        for rel in list(doc.part.rels.values()):
            if rel.reltype != reltype or rel.is_external:
                continue
            part = rel.target_part
            root_el = getattr(part, 'element', None)
            if root_el is None:
                continue
            part_name = str(part.partname)
            if part_name in seen_parts:
                continue
            seen_parts.add(part_name)

            raw: List[tuple] = []
            for note in root_el:
                if note.tag != _w(wrapper):
                    continue
                if note.get(_w('type')) in _NOTE_SKIP_TYPES:
                    continue
                note_id = note.get(_w('id'))
                label = f"{base_label} {note_id}" if note_id else base_label
                _walk_container(note, loc_type, label, raw)
            for ordinal, (p_el, l_type, detail, _section) in enumerate(raw):
                collected.append(_Located(p_el, l_type, detail,
                                          f"{part_name}#{ordinal}", ""))

    return collected


def _collect_paragraphs(doc) -> List[Tuple[object, str, str, str]]:
    """
    Every paragraph as ``(p_element, location_type, location_detail, key)``.

    The four-field view of `_locate_paragraphs`, kept for callers that do not
    need the section path.
    """
    return [(loc.element, loc.loc_type, loc.detail, loc.key)
            for loc in _locate_paragraphs(doc)]


# --------------------------------------------------------------------------
# Searching
# --------------------------------------------------------------------------

def _get_context(full_text: str, start: int, length: int, context_chars: int = 47) -> Tuple[str, str]:
    """Extract context before and after a match position, with ellipsis markers."""
    before_start = max(0, start - context_chars)
    after_end = min(len(full_text), start + length + context_chars)

    context_before = full_text[before_start:start]
    context_after = full_text[start + length:after_end]

    if before_start > 0:
        context_before = "..." + context_before
    if after_end < len(full_text):
        context_after = context_after + "..."

    return context_before, context_after


def _is_word_char(ch: str) -> bool:
    return ch.isalnum() or ch == '_'


def _is_whole_word(text: str, start: int, end: int) -> bool:
    """
    True if ``text[start:end]`` is not glued to a neighbouring word character.

    Only an edge of the match that is itself a word character needs a boundary:
    "(2022)" is a whole word inside "x(2022)" because its parenthesis edges are
    not word characters, while "2022" inside "20220" is not.
    """
    if end <= start:
        return False
    if _is_word_char(text[start]) and start > 0 and _is_word_char(text[start - 1]):
        return False
    if _is_word_char(text[end - 1]) and end < len(text) and _is_word_char(text[end]):
        return False
    return True


def _find_all_occurrences(text: str, find_text: str, case_sensitive: bool = True,
                          whole_word: bool = False) -> List[int]:
    """
    Find character offsets of every non-overlapping occurrence of find_text.

    Non-overlapping matches Word's own behaviour: "aa" occurs twice in "aaaa",
    not three times. With ``whole_word``, a candidate glued to a neighbouring
    word character is rejected and the search resumes one character later, so
    "aa" is still found in "aaa aa" (at the standalone one).
    """
    if not find_text:
        return []

    offsets = []
    haystack = text if case_sensitive else text.lower()
    needle = find_text if case_sensitive else find_text.lower()

    start = 0
    while True:
        idx = haystack.find(needle, start)
        if idx == -1:
            break
        end = idx + len(needle)
        if whole_word and not _is_whole_word(haystack, idx, end):
            start = idx + 1
            continue
        offsets.append(idx)
        start = end
    return offsets


def _normalise_rules(rules: Union[str, Rule, Sequence[Rule]],
                     case_sensitive: bool = True) -> List[Rule]:
    """
    Turn whatever the caller passed into a clean list of rules.

    A bare string is the single-term form (with an empty replacement, which is
    fine for a scan). Rules with an empty search text are dropped. Two rules
    with the same search text - compared case-insensitively when the scan is -
    would flag every one of their matches as a conflict, so that is a
    ``ValueError`` here rather than a screen full of warnings later.
    """
    if isinstance(rules, str):
        rules = [Rule(rules)]
    elif isinstance(rules, Rule):
        rules = [rules]

    cleaned: List[Rule] = []
    seen: Dict[str, str] = {}
    for rule in rules:
        if not rule.find_text:
            continue
        key = rule.find_text if case_sensitive else rule.find_text.lower()
        if key in seen:
            if seen[key] == rule.find_text:
                raise ValueError(f'Duplicate rule for "{rule.find_text}".')
            raise ValueError(
                f'Rules "{seen[key]}" and "{rule.find_text}" are the same text '
                f'when case is ignored.')
        seen[key] = rule.find_text
        cleaned.append(rule)
    return cleaned


def _is_docx_name(name: str) -> bool:
    return name.lower().endswith('.docx') and not name.startswith('~$')


def _list_files(folder_path: str, recursive: bool, wanted) -> List[str]:
    """
    Files under the folder whose name satisfies `wanted`, sorted.

    Raises the underlying OSError for the top folder. With ``recursive``,
    subfolders are walked in sorted order without following symlinks; a
    subfolder that cannot be listed is skipped rather than aborting the scan.
    """
    if not recursive:
        return sorted(
            os.path.join(folder_path, f)
            for f in os.listdir(folder_path) if wanted(f)
        )

    # Probe the top folder so a permission error there surfaces as it does in
    # the non-recursive case; os.walk would otherwise swallow it.
    os.listdir(folder_path)

    found: List[str] = []
    for root, dirs, files in os.walk(folder_path, followlinks=False):
        dirs.sort()
        for f in sorted(files):
            if wanted(f):
                found.append(os.path.join(root, f))
    return found


def _list_docx(folder_path: str, recursive: bool = False) -> List[str]:
    """Every .docx under the folder, sorted, skipping Word's ``~$`` lock files."""
    return _list_files(folder_path, recursive, _is_docx_name)


def _mark_conflicts(matches: List[Match]) -> None:
    """
    Flag matches in one paragraph whose character ranges overlap.

    Overlaps can only come from different rules ("2022" and "2022 CBC"): a
    single rule's occurrences never overlap. Neither side is applied - which
    one the user meant is not the tool's guess to make - so both are marked
    and deselected. `matches` must be sorted by offset.
    """
    for i in range(1, len(matches)):
        prev, cur = matches[i - 1], matches[i]
        prev_end = prev.char_offset + len(prev.match_text)
        if cur.char_offset < prev_end:
            prev.conflict = cur.conflict = True
            prev.is_selected = cur.is_selected = False


def scan_documents(folder_path: str, find_text: Union[str, Sequence[Rule]],
                   case_sensitive: bool = True, whole_word: bool = False,
                   recursive: bool = False) -> List[FileResult]:
    """
    Scan all .docx files in a folder for occurrences of find_text.

    `find_text` may be a single search string or a list of `Rule`s. Returns one
    FileResult per file that either matched or failed to open. Use
    `scan_documents_detailed` when the caller also needs the file count.
    """
    return scan_documents_detailed(folder_path, find_text, case_sensitive,
                                   whole_word=whole_word, recursive=recursive)[0]


def scan_documents_detailed(folder_path: str, rules: Union[str, Sequence[Rule]],
                            case_sensitive: bool = True,
                            whole_word: bool = False,
                            recursive: bool = False,
                            progress_callback=None,
                            cancel_event=None) -> Tuple[List[FileResult], dict]:
    """
    Scan a folder and also report how many .docx files were examined.

    Args:
        rules: A single search string, or a list of `Rule`s run as one set.
            Each Match records the rule it came from in `find_text` and
            `replace_text`. Matches from different rules that overlap in the
            same paragraph are flagged `conflict` and deselected.
        whole_word: Reject occurrences glued to a neighbouring word character.
        recursive: Also scan subfolders. `FileResult.file_name` is then the
            path relative to `folder_path`.
        progress_callback: Optional callable(file_name, file_index, total_files),
            called before each file so a GUI can show progress over a large set.
        cancel_event: Optional object with `is_set()`. Checked between files;
            when set the scan stops and the summary reports 'cancelled'.

    Returns ``(results, summary)`` where summary has keys 'files_scanned',
    'files_with_matches' and 'cancelled'. The caller needs 'files_scanned' to
    tell "no .docx files in this folder" apart from "no matches in any of them".

    Raises ``ValueError`` for two rules with the same search text.
    """
    summary = {'files_scanned': 0, 'files_with_matches': 0, 'cancelled': False}

    rule_list = _normalise_rules(rules, case_sensitive)
    if not rule_list:
        return [], summary

    try:
        docx_files = _list_docx(folder_path, recursive)
    except PermissionError as e:
        return [FileResult(file_path=folder_path, file_name=os.path.basename(folder_path),
                           error=f"Permission denied: {e}")], summary
    except OSError as e:
        return [FileResult(file_path=folder_path, file_name=os.path.basename(folder_path),
                           error=f"Could not read folder: {e}")], summary

    if not docx_files:
        return [], summary

    results = []

    for file_idx, file_path in enumerate(docx_files):
        if cancel_event is not None and cancel_event.is_set():
            summary['cancelled'] = True
            break

        file_name = (os.path.relpath(file_path, folder_path) if recursive
                     else os.path.basename(file_path))

        if progress_callback:
            progress_callback(file_name, file_idx, len(docx_files))

        summary['files_scanned'] += 1
        file_result = FileResult(file_path=file_path, file_name=file_name)

        # Recorded so apply_changes can tell whether the file changed underneath
        # the preview. A stale offset must never be applied blind.
        try:
            stat = os.stat(file_path)
            file_mtime, file_size = stat.st_mtime, stat.st_size
        except OSError:
            file_mtime, file_size = 0.0, 0

        try:
            doc = Document(file_path)
        except PackageNotFoundError:
            file_result.error = "File appears to be corrupted or is not a valid .docx file"
            results.append(file_result)
            continue
        except PermissionError:
            file_result.error = "Permission denied (file may be open in another application)"
            results.append(file_result)
            continue
        except Exception as e:
            file_result.error = f"Error opening file: {str(e)}"
            results.append(file_result)
            continue

        try:
            all_paragraphs = _locate_paragraphs(doc)
        except Exception as e:
            file_result.error = f"Error reading document structure: {str(e)}"
            results.append(file_result)
            continue

        for p_el, loc_type, loc_detail, para_key, section in all_paragraphs:
            full_text, _nodes = _paragraph_text_and_nodes(p_el)
            if not full_text:
                continue

            para_matches: List[Match] = []
            for rule in rule_list:
                find_text = rule.find_text
                for char_offset in _find_all_occurrences(
                        full_text, find_text, case_sensitive, whole_word):
                    context_before, context_after = _get_context(
                        full_text, char_offset, len(find_text)
                    )
                    para_matches.append(Match(
                        file_path=file_path,
                        location_type=loc_type,
                        location_detail=loc_detail,
                        paragraph_key=para_key,
                        char_offset=char_offset,
                        context_before=context_before,
                        match_text=full_text[char_offset:char_offset + len(find_text)],
                        context_after=context_after,
                        file_mtime=file_mtime,
                        file_size=file_size,
                        find_text=find_text,
                        replace_text=rule.replace_text,
                        section=section,
                    ))

            # Document order, whichever rule found them; then flag overlaps.
            para_matches.sort(key=lambda m: (m.char_offset, -len(m.match_text)))
            _mark_conflicts(para_matches)
            file_result.matches.extend(para_matches)

        if file_result.matches:
            summary['files_with_matches'] += 1
        if file_result.matches or file_result.error:
            results.append(file_result)

    return results, summary


# --------------------------------------------------------------------------
# Replacement
# --------------------------------------------------------------------------

def _set_node_text(el, text: str) -> None:
    """Set a ``w:t`` element's text in place, preserving significant whitespace."""
    el.text = text
    el.set(_XML_SPACE, "preserve")


@dataclass(frozen=True)
class _Edit:
    """One requested replacement inside a paragraph."""
    offset: int
    find_text: str
    replace_text: str


@dataclass
class _EditOutcome:
    offset: int
    applied: bool
    reason: str = ""


_REASON_OVERLAP = "overlaps another selected change"
_REASON_MOVED = "text no longer at the recorded position"
_REASON_NO_RUN = "match consists only of tabs or breaks"


def _edit_in_place(covered: List[_TextNode], offset: int, end: int,
                   replace_text: str) -> None:
    """
    Write `replace_text` over the character range ``[offset, end)``.

    The replacement goes into the first covered ``w:t``; later covered ``w:t``
    elements keep only their uncovered text. Runs are never reassigned, so
    anything else a run carries (inline images, footnote and comment
    references, field characters) survives untouched. Atomic leaves (tabs,
    breaks) inside the range are single characters, so any overlap covers the
    whole element and it is dropped - the replacement supersedes it.
    """
    placed = False
    for node in covered:
        overlap_start = max(offset, node.start)
        overlap_end = min(end, node.start + node.length)

        if node.editable:
            text = node.element.text or ""
            prefix = text[:overlap_start - node.start]
            suffix = text[overlap_end - node.start:]
            if placed:
                _set_node_text(node.element, prefix + suffix)
            else:
                _set_node_text(node.element, prefix + replace_text + suffix)
                placed = True
        else:
            parent = node.element.getparent()
            if parent is not None:
                parent.remove(node.element)


# --------------------------------------------------------------------------
# Tracked-changes mode: replacements written as Word revisions
# --------------------------------------------------------------------------

_R = _w('r')
_RPR = _w('rPr')
_INS = _w('ins')
_DEL = _w('del')
_MOVE_TO = _w('moveTo')
_DEL_TEXT = _w('delText')
_RPR_CHANGE = _w('rPrChange')

_REASON_CONFLICT = "overlaps a match from another rule"
_REASON_COMMENT_TRACKED = "revisions cannot be tracked inside a comment"
_REASON_NESTED_INS = ("inside an existing tracked insertion - accept or reject it "
                      "in Word first, or run without tracked changes")
_REASON_STRUCTURE = "unexpected document structure around the match"


@dataclass
class _TrackedContext:
    """Author, timestamp and id allocator shared by every revision in one file."""
    author: str
    date: str
    next_id: int

    def take_id(self) -> str:
        value = self.next_id
        self.next_id += 1
        return str(value)


def _default_author() -> str:
    try:
        name = getpass.getuser()
    except Exception:
        name = ""
    return name or "DocxFindReplace"


def _max_revision_id(doc) -> int:
    """
    The highest numeric ``w:id`` anywhere in the package.

    Revision ids must be unique across the document. Every part is checked
    and every ``w:id`` counts (bookmarks and comments included), which
    over-approximates harmlessly and avoids enumerating annotation types.
    """
    highest = 0
    w_id = _w('id')
    for part in doc.part.package.iter_parts():
        root = getattr(part, 'element', None)
        if root is None:
            continue
        for el in root.iter():
            value = el.get(w_id) if hasattr(el, 'get') else None
            if value is None:
                continue
            try:
                highest = max(highest, int(value))
            except (TypeError, ValueError):
                continue
    return highest


def _tracked_context_for(doc, author: Optional[str] = None) -> _TrackedContext:
    date = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    return _TrackedContext(author or _default_author(), date, _max_revision_id(doc) + 1)


def _new_t(text: str):
    t = OxmlElement('w:t')
    _set_node_text(t, text)
    return t


def _copy_rpr(rpr):
    """
    A copy of a run's ``w:rPr`` safe to attach to a new run.

    ``w:rPrChange`` (a tracked formatting change) carries its own revision id,
    so it is dropped from the copy rather than duplicated.
    """
    if rpr is None:
        return None
    copied = copy.deepcopy(rpr)
    for change in copied.findall(_RPR_CHANGE):
        copied.remove(change)
    return copied


def _new_run(rpr, children) -> object:
    """A new ``w:r`` carrying a copy of `rpr` and the given children (moved)."""
    run = OxmlElement('w:r')
    rpr_copy = _copy_rpr(rpr)
    if rpr_copy is not None:
        run.append(rpr_copy)
    for child in children:
        run.append(child)        # lxml moves an element that already has a parent
    return run


def _isolate_run_slice(node_el, a: int, b: int):
    """
    Split a run so that `node_el` - trimmed to ``text[a:b]`` when it is a
    ``w:t`` - is the only content of its run, and return that run.

    Everything else the run carried (other text, a drawing, a footnote or
    comment reference, a field character) moves into sibling runs before and
    after, each with a copy of the run's formatting. Those siblings are never
    wrapped in a revision, which is what keeps tracked mode from marking an
    image or a footnote reference as deleted.
    """
    run = node_el.getparent()
    rpr = run.find(_RPR)
    children = [c for c in run if c is not rpr]
    index = children.index(node_el)
    before, after = children[:index], children[index + 1:]

    if node_el.tag == _T:
        text = node_el.text or ""
        if a > 0:
            before.append(_new_t(text[:a]))
        if b < len(text):
            after.insert(0, _new_t(text[b:]))
        _set_node_text(node_el, text[a:b])

    if before:
        run.addprevious(_new_run(rpr, before))
    if after:
        run.addnext(_new_run(rpr, after))
    return run


def _set_revision_attrs(el, ctx: _TrackedContext) -> None:
    el.set(_w('id'), ctx.take_id())
    el.set(_w('author'), ctx.author)
    el.set(_w('date'), ctx.date)


def _wrap_in_del(run, ctx: _TrackedContext):
    """Wrap an isolated run in ``w:del``, turning its ``w:t`` into ``w:delText``."""
    wrapper = OxmlElement('w:del')
    _set_revision_attrs(wrapper, ctx)
    run.addprevious(wrapper)
    wrapper.append(run)
    for t in run.findall(_T):
        del_text = OxmlElement('w:delText')
        del_text.text = t.text
        del_text.set(_XML_SPACE, "preserve")
        run.replace(t, del_text)
    return wrapper


def _make_ins(ctx: _TrackedContext, rpr_src, text: str):
    """A ``w:ins`` holding one run of `text` formatted like `rpr_src`."""
    ins = OxmlElement('w:ins')
    _set_revision_attrs(ins, ctx)
    run = OxmlElement('w:r')
    rpr_copy = _copy_rpr(rpr_src)
    if rpr_copy is not None:
        run.append(rpr_copy)
    run.append(_new_t(text))
    ins.append(run)
    return ins


def _edit_tracked(p_el, covered: List[_TextNode], offset: int, end: int,
                  replace_text: str, ctx: _TrackedContext) -> str:
    """
    Write the replacement of ``[offset, end)`` as a Word revision.

    Each covered node is first isolated into a run of its own, then that run
    is wrapped in a tracked deletion; a tracked insertion holding the new text
    follows the last deletion, formatted like the first deleted text run. The
    traversal skips ``w:del`` and enters ``w:ins``, so the paragraph text
    afterwards reads exactly as an in-place replacement would.

    Returns "" on success, or a skip reason - in which case nothing was
    touched. Two structures are refused rather than guessed at: a match inside
    an existing tracked insertion (Word has no representation for an insertion
    nested in another), and a node whose parent is not a run.
    """
    for node in covered:
        parent = node.element.getparent()
        if parent is None or parent.tag != _R:
            return _REASON_STRUCTURE
        if _has_intervening(node.element, {_INS, _MOVE_TO}, p_el):
            return _REASON_NESTED_INS

    first_rpr = None
    last_del = None
    for node in covered:
        a = max(offset, node.start) - node.start
        b = min(end, node.start + node.length) - node.start
        run = _isolate_run_slice(node.element, a, b)
        if first_rpr is None and node.editable:
            first_rpr = run.find(_RPR)
        last_del = _wrap_in_del(run, ctx)

    if replace_text and last_del is not None:
        last_del.addnext(_make_ins(ctx, first_rpr, replace_text))
    return ""


def _apply_edits(p_el, edits: Sequence[_Edit], case_sensitive: bool = True,
                 tracked=None) -> List[_EditOutcome]:
    """
    Apply several replacements to one paragraph, each addressed by offset.

    Offsets index the paragraph text as produced by `_paragraph_text_and_nodes`
    and are applied in descending order so that earlier offsets stay valid.
    Before anything is touched, every edit is re-verified against the current
    paragraph text; an edit whose text is no longer at its offset is skipped,
    never guessed. Two requested edits whose ranges overlap are both skipped:
    applying one would silently destroy the other.

    Only the text is re-verified, deliberately: a whole-word boundary is not
    re-checked, because an adjacent edit applied a moment earlier in this same
    paragraph can legitimately have changed the neighbouring character.

    `tracked`, when given, is a `_TrackedContext`; the edit is then written as
    a Word revision instead of in place (see `_edit_tracked`).

    Returns one `_EditOutcome` per distinct edit.
    """
    outcomes: List[_EditOutcome] = []
    distinct = list(dict.fromkeys(edits))     # dedupe, preserving order

    # Overlap guard on the requested ranges, before any mutation.
    by_offset = sorted(distinct, key=lambda e: (e.offset, -len(e.find_text)))
    overlapping = set()
    for i in range(1, len(by_offset)):
        prev, cur = by_offset[i - 1], by_offset[i]
        if cur.offset < prev.offset + len(prev.find_text):
            overlapping.add(prev)
            overlapping.add(cur)

    for edit in sorted(distinct, key=lambda e: e.offset, reverse=True):
        if edit in overlapping:
            outcomes.append(_EditOutcome(edit.offset, False, _REASON_OVERLAP))
            continue
        find_len = len(edit.find_text)
        if find_len == 0:
            outcomes.append(_EditOutcome(edit.offset, False, _REASON_MOVED))
            continue

        full_text, nodes = _paragraph_text_and_nodes(p_el)
        offset, end = edit.offset, edit.offset + find_len

        # Re-verify the match is still present at this offset before touching it.
        if offset < 0 or end > len(full_text):
            outcomes.append(_EditOutcome(offset, False, _REASON_MOVED))
            continue
        actual = full_text[offset:end]
        if case_sensitive:
            same = actual == edit.find_text
        else:
            same = actual.lower() == edit.find_text.lower()
        if not same:
            outcomes.append(_EditOutcome(offset, False, _REASON_MOVED))
            continue

        covered = [n for n in nodes if n.start < end and n.start + n.length > offset]
        if not any(n.editable for n in covered):
            # Nowhere to put the replacement without inventing a run.
            outcomes.append(_EditOutcome(offset, False, _REASON_NO_RUN))
            continue

        if tracked is None:
            _edit_in_place(covered, offset, end, edit.replace_text)
            outcomes.append(_EditOutcome(offset, True))
        else:
            reason = _edit_tracked(p_el, covered, offset, end,
                                   edit.replace_text, tracked)
            outcomes.append(_EditOutcome(offset, not reason, reason))

    return outcomes


def _replace_in_paragraph(p_el, find_text: str, replace_text: str,
                          target_offsets: List[int],
                          case_sensitive: bool = True,
                          tracked=None) -> Tuple[List[int], List[int]]:
    """
    Replace specific occurrences of one search text in one paragraph.

    The single-rule form of `_apply_edits`. Returns
    ``(applied_offsets, skipped_offsets)``.
    """
    edits = [_Edit(off, find_text, replace_text) for off in target_offsets]
    outcomes = _apply_edits(p_el, edits, case_sensitive, tracked)
    applied = [o.offset for o in outcomes if o.applied]
    skipped = [o.offset for o in outcomes if not o.applied]
    return applied, skipped


def _make_backup_path(file_path: str) -> str:
    """
    Build a timestamped backup path that never overwrites an existing backup.

    e.g. ``spec.docx`` -> ``spec.docx.20260909-141530.bak``
    """
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    candidate = f"{file_path}.{stamp}.bak"
    counter = 1
    while os.path.exists(candidate):
        candidate = f"{file_path}.{stamp}-{counter}.bak"
        counter += 1
    return candidate


def _describe_staleness(file_path: str, reference: Match) -> Optional[str]:
    """
    Return a reason string if the file changed since it was scanned, else None.

    The preview records character offsets into a specific version of a file. If
    the document is edited in Word between Preview and Apply, those offsets can
    still validate against different text and quietly damage the document, so
    the file is refused rather than guessed at.

    A match with no recorded stat (constructed by hand, or a file that could not
    be stat-ed at scan time) is not checked.
    """
    if not reference.file_mtime and not reference.file_size:
        return None

    try:
        stat = os.stat(file_path)
    except OSError as e:
        return f"could not be read ({e})."

    if stat.st_size != reference.file_size:
        return ("changed on disk since the preview "
                f"(size {reference.file_size} -> {stat.st_size}).")
    if abs(stat.st_mtime - reference.file_mtime) > 0.001:
        return "was modified on disk since the preview."
    return None


def _save_atomically(doc, file_path: str) -> None:
    """
    Save to a temporary file in the same directory, then replace the original.

    python-docx streams the zip directly over the destination path, so an
    interrupted save (process killed, disk full) would otherwise leave a
    truncated, unopenable document where the original used to be.
    """
    directory = os.path.dirname(os.path.abspath(file_path))
    tmp_path = os.path.join(
        directory, f".{os.path.basename(file_path)}.{os.getpid()}.tmp"
    )
    try:
        doc.save(tmp_path)
        os.replace(tmp_path, file_path)
    except Exception:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
        raise


def apply_changes(matches: List[Match], replace_text: Optional[str] = None,
                  create_backups: bool = True, case_sensitive: bool = True,
                  tracked_changes: bool = False, author: Optional[str] = None,
                  progress_callback=None) -> dict:
    """
    Apply selected replacements to documents.

    Args:
        matches: Match objects (only those with is_selected=True are applied)
        replace_text: Override replacement text for every match. When None,
            each match's own `replace_text` (the rule it came from) is used.
        create_backups: Whether to create timestamped .bak files
        case_sensitive: Whether matching is case-sensitive
        tracked_changes: Write each replacement as a Word revision (a tracked
            deletion of the old text and insertion of the new) instead of
            editing the text in place.
        author: Revision author for tracked changes. Defaults to the current
            user's login name.
        progress_callback: Optional callable(file_name, file_index, total_files)

    Returns:
        dict with keys 'total_replaced', 'total_skipped', 'files_modified',
        'backups', 'backup_map', 'skip_reasons' and 'errors'. Every selected
        match is accounted for as either replaced or skipped, each Match's
        `applied` attribute is set, and a skipped match carries a
        `skip_reason`. 'skip_reasons' counts skipped matches by reason.
    """
    selected = [m for m in matches if m.is_selected]
    for m in selected:
        m.applied = None
        m.skip_reason = ""

    result = {'total_replaced': 0, 'total_skipped': 0, 'files_modified': 0,
              'backups': [], 'backup_map': {}, 'skip_reasons': {}, 'errors': []}
    if not selected:
        return result

    files_dict: Dict[str, List[Match]] = {}
    for match in selected:
        files_dict.setdefault(match.file_path, []).append(match)

    if tracked_changes and not author:
        author = _default_author()

    errors: List[str] = result['errors']
    total_files = len(files_dict)

    def skip_all(file_matches: List[Match], reason: str) -> None:
        for m in file_matches:
            m.applied = False
            m.skip_reason = reason

    for file_idx, (file_path, file_matches) in enumerate(files_dict.items()):
        file_name = os.path.basename(file_path)

        if progress_callback:
            progress_callback(file_name, file_idx, total_files)

        # Accounting runs in the `finally` below so that every early exit from
        # this block - an unreadable file, a backup that cannot be written, a
        # failed save - still reports its matches as skipped. A selected change
        # must never disappear from the totals.
        try:
            stale = _describe_staleness(file_path, file_matches[0])
            if stale:
                errors.append(f"{file_name}: {stale} Nothing was changed in this "
                              f"file - re-run Preview Changes.")
                skip_all(file_matches, "file changed since the preview")
                continue

            try:
                doc = Document(file_path)
            except Exception as e:
                errors.append(f"Could not open {file_name}: {e}")
                skip_all(file_matches, "file could not be opened")
                continue

            try:
                all_paragraphs = _collect_paragraphs(doc)
            except Exception as e:
                errors.append(f"Could not read structure of {file_name}: {e}")
                skip_all(file_matches, "document structure could not be read")
                continue

            tracked = _tracked_context_for(doc, author) if tracked_changes else None

            para_lookup = {key: p_el for p_el, _t, _d, key in all_paragraphs}

            by_paragraph: Dict[str, List[Match]] = {}
            for match in file_matches:
                if match.conflict:
                    # Flagged at scan time; never applied, whatever the
                    # checkbox says.
                    match.applied = False
                    match.skip_reason = _REASON_CONFLICT
                    continue
                if tracked is not None and match.location_type == 'comment':
                    # Word does not track revisions inside comment text, so a
                    # revision there would be a guess about how Word shows it.
                    match.applied = False
                    match.skip_reason = _REASON_COMMENT_TRACKED
                    continue
                by_paragraph.setdefault(match.paragraph_key, []).append(match)

            file_replaced = 0

            for para_key, para_matches in by_paragraph.items():
                p_el = para_lookup.get(para_key)
                if p_el is None:
                    skip_all(para_matches, "paragraph not found in the document")
                    continue

                edits = [
                    _Edit(m.char_offset, m.effective_find_text,
                          replace_text if replace_text is not None else m.replace_text)
                    for m in para_matches
                ]
                outcomes = {o.offset: o for o in
                            _apply_edits(p_el, edits, case_sensitive, tracked)}
                for m in para_matches:
                    outcome = outcomes.get(m.char_offset)
                    if outcome is not None and outcome.applied:
                        m.applied = True
                    else:
                        m.applied = False
                        m.skip_reason = outcome.reason if outcome else _REASON_MOVED
                file_replaced += sum(1 for o in outcomes.values() if o.applied)

            if file_replaced > 0:
                backup_path = None
                if create_backups:
                    backup_path = _make_backup_path(file_path)
                    try:
                        shutil.copy2(file_path, backup_path)
                    except Exception as e:
                        errors.append(f"Could not create backup for {file_name}: {e} "
                                      f"(file left unchanged)")
                        skip_all(file_matches, "backup could not be written")
                        continue

                try:
                    _save_atomically(doc, file_path)
                    result['files_modified'] += 1
                    result['total_replaced'] += file_replaced
                    if backup_path:
                        result['backups'].append(backup_path)
                        result['backup_map'][file_path] = backup_path
                except PermissionError:
                    errors.append(f"Permission denied saving {file_name} (file may be open)")
                    skip_all(file_matches, "permission denied saving the file")
                except Exception as e:
                    errors.append(f"Error saving {file_name}: {e}")
                    skip_all(file_matches, "file could not be saved")
        finally:
            for m in file_matches:
                if not m.applied:
                    result['total_skipped'] += 1
                    reason = m.skip_reason or "not applied"
                    result['skip_reasons'][reason] = result['skip_reasons'].get(reason, 0) + 1

    if progress_callback:
        progress_callback("Done", total_files - 1 if total_files else 0, total_files)

    return result


# --------------------------------------------------------------------------
# Change log
# --------------------------------------------------------------------------

CHANGE_LOG_COLUMNS = ("file", "location", "section", "find", "replace",
                      "before", "match", "after", "status", "backup")


def _match_status(match: Match) -> str:
    if not match.is_selected:
        return "not selected"
    if match.applied:
        return "applied"
    return f"skipped: {match.skip_reason or 'not applied'}"


def write_change_log(path: str, matches: Sequence[Match],
                     backup_map: Optional[Dict[str, str]] = None) -> str:
    """
    Write a CSV record of a run: one row per match, applied or not.

    Written as UTF-8 with a byte-order mark so Excel opens it correctly, with
    the columns in `CHANGE_LOG_COLUMNS`. The status column reads ``applied``,
    ``skipped: <reason>`` or ``not selected``. Returns `path`.
    """
    backup_map = backup_map or {}
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(CHANGE_LOG_COLUMNS)
        for m in matches:
            writer.writerow([
                m.file_path,
                f"{m.location_type}: {m.location_detail}",
                m.section,
                m.effective_find_text,
                m.replace_text,
                m.context_before,
                m.match_text,
                m.context_after,
                _match_status(m),
                backup_map.get(m.file_path, ""),
            ])
    return path


# --------------------------------------------------------------------------
# Backups: discovery and restore
# --------------------------------------------------------------------------

_BACKUP_RE = re.compile(
    r'^(?P<doc>.+\.docx)\.(?P<stamp>\d{8}-\d{6})(?:-(?P<n>\d+))?\.bak$',
    re.IGNORECASE,
)


class Backup(NamedTuple):
    """A timestamped backup written by `apply_changes` and the document it is of."""
    document_path: str
    backup_path: str
    timestamp: datetime
    sequence: int          # the -N collision suffix, 0 when absent


def _parse_backup(path: str) -> Optional[Backup]:
    m = _BACKUP_RE.match(os.path.basename(path))
    if not m:
        return None
    try:
        stamp = datetime.strptime(m.group('stamp'), "%Y%m%d-%H%M%S")
    except ValueError:
        return None
    document = os.path.join(os.path.dirname(path), m.group('doc'))
    return Backup(document, path, stamp, int(m.group('n') or 0))


def find_backups(folder_path: str, recursive: bool = False) -> List[Backup]:
    """
    Every backup under the folder, grouped by document, newest first.

    Only files matching the name pattern `_make_backup_path` writes are
    returned; an unrelated ``.bak`` is ignored. Raises the folder's OSError.
    """
    names = _list_files(folder_path, recursive, lambda f: f.lower().endswith('.bak'))
    backups = [b for b in (_parse_backup(p) for p in names) if b is not None]
    backups.sort(key=lambda b: (b.document_path.lower(), b.document_path,
                                -b.timestamp.timestamp(), -b.sequence))
    return backups


def restore_backups(backups: Sequence[Backup], progress_callback=None) -> dict:
    """
    Copy each backup back over its document.

    The copy goes to a temporary file beside the document and is swapped in
    with ``os.replace``, so an interrupted restore cannot leave a truncated
    document. The backup file is left in place. Two backups of the same
    document in one call is a ``ValueError`` raised before anything is
    touched. Returns ``{'restored': n, 'restored_paths': [...], 'errors': [...]}``.
    """
    seen = set()
    for b in backups:
        key = os.path.normcase(os.path.abspath(b.document_path))
        if key in seen:
            raise ValueError(f"More than one backup selected for {b.document_path}")
        seen.add(key)

    result = {'restored': 0, 'restored_paths': [], 'errors': []}
    for idx, b in enumerate(backups):
        name = os.path.basename(b.document_path)
        if progress_callback:
            progress_callback(name, idx, len(backups))
        directory = os.path.dirname(os.path.abspath(b.document_path))
        tmp_path = os.path.join(directory, f".{name}.{os.getpid()}.tmp")
        try:
            shutil.copy2(b.backup_path, tmp_path)
            os.replace(tmp_path, b.document_path)
            result['restored'] += 1
            result['restored_paths'].append(b.document_path)
        except Exception as e:
            result['errors'].append(f"Could not restore {name}: {e}")
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass
    return result
