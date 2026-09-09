# Word Document Batch Find & Replace

A Python GUI application for performing batch find-and-replace operations across multiple Word documents (.docx) with a preview-before-commit workflow.

## Use Case

Updating technical specifications and building code references across multiple Word documents — for example, replacing "2022 CBC" with "2025 CBC" while being able to skip historical references like "designed in 2022."

## Features

- **Preview before commit** — Scan all documents and review every match with surrounding context before applying any changes
- **Selective replacement** — Check/uncheck individual matches or entire files; only approved changes are applied
- **Cross-run matching** — Finds text even when Word has split it across multiple formatting runs (a common issue with `python-docx`)
- **Hyperlink-aware** — Text inside hyperlinks is searched and replaced at the correct offset, and the link relationship survives
- **Accurate match counts** — A physical paragraph is visited exactly once, so merged table cells and headers shared between sections are counted and replaced once, not once per section or spanned column
- **Non-destructive replacement** — Inline images, footnote and comment references, and field characters sharing a run with matched text are left intact
- **Formatting preservation** — Replacement text inherits the formatting of the run it lands in; surrounding runs keep bold, italic, font and color
- **Timestamped backups** — Optional `.bak` files that never overwrite each other, so the true original stays recoverable across repeated runs
- **Atomic saves** — Documents are written to a temporary file and swapped into place, so an interrupted save cannot truncate the original
- **Full accounting** — Every selected change is reported as either applied or skipped; a change that could not be applied is never silently dropped
- **Case sensitivity toggle** — Search with or without case sensitivity
- **Error handling** — Gracefully skips corrupted, locked, or permission-denied files with clear error messages

## Requirements

- Python 3.8+
- `python-docx` 1.1.0 or newer (see [Why the version floor matters](#why-the-version-floor-matters))
- Windows (tested), macOS/Linux (should work with tkinter installed)

## Installation

```bash
cd docx-find-and-replace
pip install -r requirements.txt
```

To run the test suite as well:

```bash
pip install -r requirements-dev.txt
```

## Usage

```bash
python main.py
```

1. **Select folder** — Click "Browse..." and choose the folder containing your `.docx` files (non-recursive, skips temp files like `~$*.docx`)
2. **Enter find/replace text** — Type the text to search for and the replacement text
3. **Configure options** — Toggle case sensitivity and backup creation
4. **Preview** — Click "Preview Changes" to scan all documents. Results appear in a tree view:
   - File-level nodes show total match count
   - Each match shows its location (body/table/header/footer) and surrounding context with the match highlighted in brackets
5. **Select/deselect** — Click any match to toggle it. Click a file node to toggle all matches in that file. Use "Select All" / "Deselect All" buttons for bulk operations.
6. **Apply** — Click "Apply Selected Changes." A confirmation dialog shows the summary before proceeding. The completion dialog reports replacements made, files modified, backups created, and — if any selected change could not be applied — a warning saying how many were skipped.

## File Structure

```
docx-find-and-replace/
├── main.py                 # GUI application (tkinter)
├── document_processor.py   # Core scan and replace logic
├── requirements.txt        # Runtime dependencies
├── requirements-dev.txt    # Test dependencies
├── pytest.ini              # Test configuration
├── tests/
│   ├── conftest.py         # Fixture builders (hyperlinks, merged cells, images, ...)
│   ├── test_traversal.py   # Text extraction and paragraph identity
│   ├── test_replace.py     # Paragraph-level replacement
│   ├── test_search.py      # Occurrence finding, context, folder scanning
│   └── test_apply.py       # End-to-end scan → apply, backups, reporting
└── README.md               # This file
```

## Document Coverage

Scanned and replaced:

- Body paragraphs
- Tables, including tables nested inside table cells
- Headers and footers (default, first-page, and even-page)
- Text boxes
- Text inside hyperlinks
- Text inside tracked insertions (`w:ins`), smart tags, and content controls

Deliberately not touched:

- Text inside tracked deletions (`w:del`) — this text is already marked for removal
- Cached field results (`w:fldSimple`) — Word regenerates these on field update, so editing them has no lasting effect

See [Known Limitations](#known-limitations) for content that is not reached at all.

## How It Works

### Cross-run matching

Word internally stores paragraph text as a sequence of "runs" — segments that share the same formatting. A single phrase like "2022 CBC" can be split across multiple runs due to spell-check history, partial edits, or even just opening and re-saving. For example:

```
Run 1: "Reference the 20"  (bold)
Run 2: "22 CBC"            (bold + red)
Run 3: " for compliance."  (normal)
```

A naive `run.text.replace()` approach would miss this entirely.

### One traversal, one source of truth

The tool walks the paragraph's XML once, and that single pass produces both the paragraph's text and the map from each character offset back to the XML element that holds it. Because they are built together, they cannot disagree.

This matters because `python-docx`'s convenience APIs disagree with each other. `Paragraph.text` is assembled from `w:r | w:hyperlink`, while `Paragraph.runs` returns only direct `w:r` children. Searching one and editing the other puts every replacement in a hyperlink-containing paragraph at the wrong offset — silently deleting unrelated prose while leaving the actual target untouched.

### Paragraph identity

The scan and the apply phase open each file separately, so a paragraph has to be identifiable across two independent loads. Each paragraph is keyed by its **part name plus its ordinal position in a deterministic walk of that part** (for example `/word/document.xml#42`). Python object identity would not survive the reload, and a positional counter over `python-docx` objects double-counts:

- `row.cells` yields the same merged cell once per spanned grid column
- `section.header` returns the same inherited header for every section linked to the previous one

Walking the XML visits each physical `w:tc` and each header *part* exactly once, so a merged cell spanning three columns is one match, and a header shared by three sections is one match.

### In-place text edits

Replacements are written directly into `w:t` element text. Assigning to `Run.text` would call `clear_content()`, which strips *every* child of the run — destroying inline images, footnote references, comment anchors and field characters that happen to share a run with the matched text.

### Why the version floor matters

`python-docx` changed the behaviour of `Paragraph.text` between 0.8.x and 1.1: older versions exclude hyperlink text, newer versions include it. Code that depends on that property behaves differently depending on which version `pip` happened to resolve. This tool reads the XML directly rather than depending on that property, but the floor is pinned at 1.1.0 anyway so that the APIs it does use (`part.related_parts`, `section._sectPr`, part element access) are guaranteed present.

## Testing

```bash
pip install -r requirements-dev.txt
pytest
```

The suite builds real `.docx` files on disk for each structure the replacement logic has to get right — hyperlinks, merged cells, headers linked across sections, inline images inside a spanned run, nested tables, text boxes in both VML and `mc:AlternateContent` form, and tracked changes — then asserts on the resulting documents. Each test named after a corruption case asserts the document is left *correct*, not merely that the call returned without raising.

## Known Limitations

- **Non-recursive** — Only processes `.docx` files in the selected folder, not subfolders
- **No regex** — Find text is matched literally (case-sensitive or case-insensitive)
- **No undo** — Use the backup feature; there is no built-in undo. Backups are timestamped (`spec.docx.20260909-141530.bak`) and are only created for files that are actually modified. Rename one back to `.docx` to restore it.
- **Not scanned: footnotes, endnotes and comments** — These live in separate document parts that are not yet traversed
- **Text box fallback copies** — Word stores a text box twice, as a modern `mc:Choice` copy and a legacy `mc:Fallback` copy. Only the `mc:Choice` copy is replaced, which is what current versions of Word render; the fallback copy retains the old text until Word next re-saves the file
- **Field results are not replaced** — Text produced by a field (cross-references, a table of contents) is regenerated by Word from the field code, so replacing the cached result would not persist
- **Preview can go stale** — If a document is edited in Word between clicking Preview and clicking Apply, the recorded offsets may no longer be valid. Affected changes are skipped rather than misapplied, and the completion dialog reports how many were skipped, but re-running the preview is the reliable move.
- **Single-pass matching** — Replacement text is not re-scanned, so replacing "2022" with "2022 and 2025" applies exactly once per original occurrence

## Copyright Notice

**Copyright © 2025 Abraham Borg. All Rights Reserved.**

This software and associated documentation files (the "Software") are the proprietary property of Abraham Borg. 

**Unauthorized copying, modification, distribution, or use of this Software, via any medium, is strictly prohibited without express written permission from the copyright holder.**

This Software is provided for review and reference purposes only. No license or right to use, copy, modify, or distribute this Software for any purpose, commercial or non-commercial, is granted.
