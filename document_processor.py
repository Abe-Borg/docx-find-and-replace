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
headers and footers. Text inside hyperlinks, tracked insertions (``w:ins``),
smart tags and content controls is included. See README "Known Limitations" for
what is deliberately not covered.
"""

import os
import shutil
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional, Tuple

from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from docx.oxml.ns import qn


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


@dataclass
class Match:
    """Represents a single find-text occurrence in a document."""
    file_path: str
    location_type: str          # 'body', 'table', 'header', 'footer'
    location_detail: str        # e.g., 'Paragraph 5', 'Table 2, Row 1, Cell 3'
    paragraph_key: str          # Structural key, stable across document loads
    char_offset: int            # Character offset within the paragraph's full text
    context_before: str         # ~50 chars before match
    match_text: str             # The matched text
    context_after: str          # ~50 chars after match
    is_selected: bool = True    # User's checkbox state
    applied: Optional[bool] = None   # Set by apply_changes(): True, False, or None

    @property
    def display_context(self) -> str:
        """Format the context for display in the GUI (single line)."""
        before = self.context_before.replace("\n", "⏎ ").replace("\t", " ")
        after = self.context_after.replace("\n", " ⏎").replace("\t", " ")
        match = self.match_text.replace("\n", "⏎").replace("\t", " ")
        return f"{before}[{match}]{after}"


@dataclass
class FileResult:
    """All matches for a single file."""
    file_path: str
    file_name: str
    matches: List[Match] = field(default_factory=list)
    error: Optional[str] = None

    @property
    def match_count(self) -> int:
        return len(self.matches)

    @property
    def selected_count(self) -> int:
        return sum(1 for m in self.matches if m.is_selected)


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


def _walk_textboxes(p_el, loc_type: str, para_detail: str, out: List[tuple]) -> None:
    """Collect paragraphs of any text box anchored in this paragraph."""
    for txbx in p_el.iter(_TXBX_CONTENT):
        # Skip the legacy mc:Fallback copy of a text box; Word renders mc:Choice.
        if _has_intervening(txbx, {_FALLBACK}, p_el):
            continue
        # A nested text box is reached through its own paragraph's walk.
        if _has_intervening(txbx, {_TXBX_CONTENT}, p_el):
            continue
        _walk_container(txbx, loc_type, f"{para_detail}, Text Box", out)


def _walk_table(tbl_el, loc_type: str, prefix: str, table_no: int, out: List[tuple]) -> None:
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
            _walk_container(cell, cell_type, cell_prefix, out)


def _walk_container(el, loc_type: str, prefix: str, out: List[tuple]) -> None:
    """
    Walk a block-level container (``w:body``, ``w:hdr``, ``w:ftr``, ``w:tc``,
    ``w:txbxContent``) in document order, appending
    ``(p_element, location_type, location_detail)`` tuples to `out`.
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
            out.append((child, loc_type, detail))
            _walk_textboxes(child, loc_type, detail, out)
        elif tag == _TBL:
            table_no += 1
            _walk_table(child, loc_type, prefix, table_no, out)
        elif tag == _SDT:
            content = child.find(_SDT_CONTENT)
            if content is not None:
                _walk_container(content, loc_type, prefix, out)
        elif tag == _ALT_CONTENT:
            choice = child.find(_CHOICE)
            if choice is not None:
                _walk_container(choice, loc_type, prefix, out)


def _collect_paragraphs(doc) -> List[Tuple[object, str, str, str]]:
    """
    Collect every paragraph in the document exactly once.

    Returns ``(p_element, location_type, location_detail, paragraph_key)``.

    `paragraph_key` is ``"<part name>#<ordinal>"``. The ordinal is the position
    in this deterministic XML walk of that part, so the same physical paragraph
    receives the same key on every load of an unchanged file.
    """
    collected: List[Tuple[object, str, str, str]] = []

    def add_part(part_name: str, root_el, loc_type: str, label: str) -> None:
        raw: List[tuple] = []
        _walk_container(root_el, loc_type, label, raw)
        for ordinal, (p_el, l_type, detail) in enumerate(raw):
            collected.append((p_el, l_type, detail, f"{part_name}#{ordinal}"))

    # --- Body (includes tables, nested tables and text boxes)
    add_part(str(doc.part.partname), doc.element.body, 'body', "")

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
    seen_parts = set()

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

    return collected


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


def _find_all_occurrences(text: str, find_text: str, case_sensitive: bool = True) -> List[int]:
    """
    Find character offsets of every non-overlapping occurrence of find_text.

    Non-overlapping matches Word's own behaviour: "aa" occurs twice in "aaaa",
    not three times.
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
        offsets.append(idx)
        start = idx + len(needle)
    return offsets


def scan_documents(folder_path: str, find_text: str, case_sensitive: bool = True) -> List[FileResult]:
    """
    Scan all .docx files in a folder for occurrences of find_text.

    Returns one FileResult per file that either matched or failed to open.
    Use `scan_summary` for the number of files actually examined.
    """
    return scan_documents_detailed(folder_path, find_text, case_sensitive)[0]


def scan_documents_detailed(folder_path: str, find_text: str,
                            case_sensitive: bool = True) -> Tuple[List[FileResult], dict]:
    """
    Scan a folder and also report how many .docx files were examined.

    Returns ``(results, summary)`` where summary has keys 'files_scanned' and
    'files_with_matches'. The caller needs 'files_scanned' to tell "no .docx
    files in this folder" apart from "no matches in any of them".
    """
    summary = {'files_scanned': 0, 'files_with_matches': 0}

    if not find_text:
        return [], summary

    try:
        docx_files = sorted([
            os.path.join(folder_path, f)
            for f in os.listdir(folder_path)
            if f.lower().endswith('.docx') and not f.startswith('~$')
        ])
    except PermissionError as e:
        return [FileResult(file_path=folder_path, file_name=os.path.basename(folder_path),
                           error=f"Permission denied: {e}")], summary
    except OSError as e:
        return [FileResult(file_path=folder_path, file_name=os.path.basename(folder_path),
                           error=f"Could not read folder: {e}")], summary

    summary['files_scanned'] = len(docx_files)
    if not docx_files:
        return [], summary

    results = []

    for file_path in docx_files:
        file_result = FileResult(
            file_path=file_path,
            file_name=os.path.basename(file_path)
        )

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
            all_paragraphs = _collect_paragraphs(doc)
        except Exception as e:
            file_result.error = f"Error reading document structure: {str(e)}"
            results.append(file_result)
            continue

        for p_el, loc_type, loc_detail, para_key in all_paragraphs:
            full_text, _nodes = _paragraph_text_and_nodes(p_el)
            if not full_text:
                continue

            for char_offset in _find_all_occurrences(full_text, find_text, case_sensitive):
                context_before, context_after = _get_context(
                    full_text, char_offset, len(find_text)
                )
                file_result.matches.append(Match(
                    file_path=file_path,
                    location_type=loc_type,
                    location_detail=loc_detail,
                    paragraph_key=para_key,
                    char_offset=char_offset,
                    context_before=context_before,
                    match_text=full_text[char_offset:char_offset + len(find_text)],
                    context_after=context_after,
                ))

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


def _replace_in_paragraph(p_el, find_text: str, replace_text: str,
                          target_offsets: List[int],
                          case_sensitive: bool = True) -> Tuple[List[int], List[int]]:
    """
    Replace specific occurrences of find_text in one paragraph.

    Occurrences are addressed by character offset into the paragraph text as
    produced by `_paragraph_text_and_nodes`, and are applied in descending
    offset order so that earlier offsets stay valid.

    Text is written directly into ``w:t`` elements. Runs are never reassigned,
    so anything else a run carries (inline images, footnote and comment
    references, field characters) survives untouched.

    Returns ``(applied_offsets, skipped_offsets)``.
    """
    applied: List[int] = []
    skipped: List[int] = []
    find_len = len(find_text)

    if find_len == 0:
        return applied, list(target_offsets)

    for offset in sorted(set(target_offsets), reverse=True):
        full_text, nodes = _paragraph_text_and_nodes(p_el)
        end = offset + find_len

        # Re-verify the match is still present at this offset before touching it.
        if offset < 0 or end > len(full_text):
            skipped.append(offset)
            continue
        actual = full_text[offset:end]
        if case_sensitive:
            if actual != find_text:
                skipped.append(offset)
                continue
        elif actual.lower() != find_text.lower():
            skipped.append(offset)
            continue

        covered = [n for n in nodes if n.start < end and n.start + n.length > offset]
        if not any(n.editable for n in covered):
            # Match consists only of atomic leaves (tabs/breaks); nowhere to put
            # the replacement without inventing a run. Leave it alone.
            skipped.append(offset)
            continue

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
                # Atomic leaves are single characters, so any overlap covers the
                # whole element: drop it, the replacement text supersedes it.
                parent = node.element.getparent()
                if parent is not None:
                    parent.remove(node.element)

        applied.append(offset)

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


def apply_changes(matches: List[Match], replace_text: str,
                  create_backups: bool = True, case_sensitive: bool = True,
                  progress_callback=None) -> dict:
    """
    Apply selected replacements to documents.

    Args:
        matches: Match objects (only those with is_selected=True are applied)
        replace_text: The replacement text
        create_backups: Whether to create timestamped .bak files
        case_sensitive: Whether matching is case-sensitive
        progress_callback: Optional callable(file_name, file_index, total_files)

    Returns:
        dict with keys 'total_replaced', 'total_skipped', 'files_modified',
        'backups', and 'errors'. Every selected match is accounted for as
        either replaced or skipped, and each Match's `applied` attribute is set.
    """
    selected = [m for m in matches if m.is_selected]
    for m in selected:
        m.applied = None

    if not selected:
        return {'total_replaced': 0, 'total_skipped': 0, 'files_modified': 0,
                'backups': [], 'errors': []}

    files_dict: Dict[str, List[Match]] = {}
    for match in selected:
        files_dict.setdefault(match.file_path, []).append(match)

    total_replaced = 0
    total_skipped = 0
    files_modified = 0
    backups: List[str] = []
    errors: List[str] = []
    total_files = len(files_dict)

    for file_idx, (file_path, file_matches) in enumerate(files_dict.items()):
        file_name = os.path.basename(file_path)

        if progress_callback:
            progress_callback(file_name, file_idx, total_files)

        try:
            doc = Document(file_path)
        except Exception as e:
            errors.append(f"Could not open {file_name}: {e}")
            for m in file_matches:
                m.applied = False
            total_skipped += len(file_matches)
            continue

        try:
            all_paragraphs = _collect_paragraphs(doc)
        except Exception as e:
            errors.append(f"Could not read structure of {file_name}: {e}")
            for m in file_matches:
                m.applied = False
            total_skipped += len(file_matches)
            continue

        para_lookup = {key: p_el for p_el, _t, _d, key in all_paragraphs}

        # Group this file's matches by the paragraph they belong to.
        by_paragraph: Dict[str, List[Match]] = {}
        for match in file_matches:
            by_paragraph.setdefault(match.paragraph_key, []).append(match)

        file_replaced = 0

        for para_key, para_matches in by_paragraph.items():
            p_el = para_lookup.get(para_key)
            if p_el is None:
                # The document changed between scanning and applying.
                for m in para_matches:
                    m.applied = False
                continue

            find_text = para_matches[0].match_text
            applied_offsets, _skipped = _replace_in_paragraph(
                p_el, find_text, replace_text,
                [m.char_offset for m in para_matches],
                case_sensitive,
            )
            applied_set = set(applied_offsets)
            for m in para_matches:
                m.applied = m.char_offset in applied_set
            file_replaced += len(applied_set)

        if file_replaced > 0:
            # Back up only once we know the file is genuinely about to change.
            backup_path = None
            if create_backups:
                backup_path = _make_backup_path(file_path)
                try:
                    shutil.copy2(file_path, backup_path)
                except Exception as e:
                    errors.append(f"Could not create backup for {file_name}: {e} "
                                  f"(file left unchanged)")
                    for m in file_matches:
                        m.applied = False
                    continue

            try:
                _save_atomically(doc, file_path)
                files_modified += 1
                total_replaced += file_replaced
                if backup_path:
                    backups.append(backup_path)
            except PermissionError:
                errors.append(f"Permission denied saving {file_name} (file may be open)")
                for m in file_matches:
                    m.applied = False
            except Exception as e:
                errors.append(f"Error saving {file_name}: {e}")
                for m in file_matches:
                    m.applied = False

        total_skipped += sum(1 for m in file_matches if not m.applied)

    if progress_callback:
        progress_callback("Done", total_files - 1 if total_files else 0, total_files)

    return {
        'total_replaced': total_replaced,
        'total_skipped': total_skipped,
        'files_modified': files_modified,
        'backups': backups,
        'errors': errors,
    }
