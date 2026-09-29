# Word Document Batch Find & Replace

A Python GUI application for performing batch find-and-replace operations across multiple Word documents (.docx) with a preview-before-commit workflow.

## Use Case

Updating technical specifications and building code references across multiple Word documents — for example, replacing "2022 CBC" with "2025 CBC" while being able to skip historical references like "designed in 2022."

A code-cycle update is rarely one replacement. It is a dozen — "2022 CBC" → "2025 CBC", "2022 CFC" → "2025 CFC", "NFPA 13, 2022 edition" → "NFPA 13, 2025 edition", and the section numbers that moved — run across a whole specification folder, and often reviewed by an engineer of record before it is accepted. The tool is built around that: a saved **rule set** run as one preview, a **change log** for the QA record, **restore** from the backups it took, and the option to write every replacement as a Word **tracked change** so the reviewer accepts or rejects it in Word.

## Features

- **Preview before commit** — Scan all documents and review every match with surrounding context before applying any changes
- **Selective replacement** — Check/uncheck individual matches or entire files; only approved changes are applied
- **Rule sets** — Run any number of find/replace rules in one preview and one apply. Save a set to a plain CSV and load it again next code cycle. Matches from different rules that overlap in the same text are flagged and never applied
- **Tracked-changes mode** — Write each replacement as a Word revision (a tracked deletion of the old text and insertion of the new) so it can be reviewed, accepted or rejected in Word instead of editing the text in place
- **Section context** — Each match shows the headings above it ("GENERAL > REFERENCES"), not just a paragraph number
- **Change log** — Every apply writes a CSV beside the documents listing each match, what it became, its status and the backup that was taken
- **Restore from backup** — Put documents back from the timestamped backups the tool wrote, from inside the tool
- **Whole-word matching** — Optionally refuse matches glued to a neighbouring letter or digit, so "2022" does not match inside "20220"
- **Subfolders** — Optionally scan the whole folder tree
- **Remembers itself** — The folder, rules and options come back the way they were left
- **Cross-run matching** — Finds text even when Word has split it across multiple formatting runs (a common issue with `python-docx`)
- **Hyperlink-aware** — Text inside hyperlinks is searched and replaced at the correct offset, and the link relationship survives
- **Accurate match counts** — A physical paragraph is visited exactly once, so merged table cells and headers shared between sections are counted and replaced once, not once per section or spanned column
- **Non-destructive replacement** — Inline images, footnote and comment references, and field characters sharing a run with matched text are left intact
- **Formatting preservation** — Replacement text inherits the formatting of the run it lands in; surrounding runs keep bold, italic, font and color
- **Timestamped backups** — Optional `.bak` files that never overwrite each other, so the true original stays recoverable across repeated runs
- **Atomic saves** — Documents are written to a temporary file and swapped into place, so an interrupted save cannot truncate the original
- **Full accounting** — Every selected change is reported as either applied or skipped; a change that could not be applied is never silently dropped
- **Stale-preview protection** — A document edited in Word between Preview and Apply is refused rather than edited with offsets that no longer describe it
- **Cancellable scan** — Large batches report progress file by file and can be stopped without waiting for the whole folder
- **Case sensitivity toggle** — Search with or without case sensitivity
- **Error handling** — Gracefully skips corrupted, locked, or permission-denied files with clear error messages

## Requirements

- Python 3.8+
- `python-docx` 1.1.0 or newer (see [Why the version floor matters](#why-the-version-floor-matters))
- Windows (tested), macOS/Linux (should work with tkinter installed)

## Installation

### Installer (no Python needed)

Download `DocxFindReplace-Setup-<version>.exe` from the
[Releases](https://github.com/Abe-Borg/docx-find-and-replace/releases) page and
run it. It installs for the current user only, so it needs no administrator
rights on a managed machine, and adds a Start Menu entry plus an uninstaller.

A single-file `DocxFindReplace-Portable.exe` is published alongside it for
running from a USB stick or a network share without installing anything. It
starts more slowly than the installed build, because it unpacks itself on each
launch, and is likelier to be held up by antivirus scanning.

Neither executable is code-signed, so Windows SmartScreen will show a
"Windows protected your PC" warning the first time. Choose **More info** →
**Run anyway**.

### From source

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

1. **Select folder** — Click "Browse..." and choose the folder containing your `.docx` files. Tick **Include subfolders** to scan the whole tree. Word's `~$*.docx` lock files are skipped either way.
2. **Build the rule set** — Type a find text and its replacement and click **Add / Update**. Repeat for every rule. Selecting a rule in the list loads it back into the boxes so it can be changed, and **Add / Update** then replaces that rule whichever box was edited; **Remove** drops the selected rules. **Load set...** and **Save set...** read and write the list as a CSV file (see [Rule-set files](#rule-set-files)). For a single replacement there is no need to click Add: Preview picks up whatever is still typed in the boxes.
3. **Configure options** — **Case sensitive** and **Whole word only** decide what counts as a match.
4. **Preview** — Click "Preview Changes" to scan all documents. The status bar names each file as it is scanned, and the button becomes "Cancel Scan" so a large batch can be stopped early. Results appear in a tree view:
   - File-level nodes show total match count
   - Each match shows its location (body/table/header/footer/footnote/endnote/comment) and surrounding context with the match highlighted in brackets, the rule that found it in the **Rule** column, and the headings above it in the **Section** column
   - A match marked ⚠ overlaps a match from another rule (for example "2022" and "2022 CBC" on the same text). Neither is applied; change the rules and preview again to resolve it
5. **Select/deselect** — Click any match to toggle it. Click a file node to toggle all matches in that file. Use "Select All" / "Deselect All" buttons for bulk operations. Clicking the expand/collapse arrow only expands; it does not toggle anything. Double-click a row to open that document in Word.

   Changing the folder, the rule set, or the case or whole-word setting discards the results on screen — they describe offsets found under the old settings, so the preview has to be re-run. Typing in the find/replace boxes does not, until the rule is added.
6. **Choose how to apply** — **Create backups** keeps a timestamped copy of every file that changes. **Write change log (CSV)** records the run beside the documents. **Write as tracked changes** writes each replacement as a Word revision under the **Author** name (the Windows user name when left blank) instead of editing the text in place — see [Tracked-changes mode](#tracked-changes-mode).
7. **Apply** — Click "Apply Selected Changes." A confirmation dialog lists the rules that were scanned, the mode, and whether backups will be taken. The completion dialog reports replacements made, files modified, backups created, the change log's path, and — if any selected change could not be applied — how many were skipped and why.
8. **Undo if needed** — **Restore backups...** lists every timestamped backup under the folder with the newest one of each document pre-ticked. Restoring copies the backup back over the document and leaves the backup file in place; edits made to the document since the backup was taken are lost.

The folder, rule set and options are saved when the window closes and restored when it opens, in `%APPDATA%\DocxFindReplace\settings.json`. Delete that file to start clean.

### Rule-set files

A rule set is a CSV with a `find,replace` header and one rule per row:

```
find,replace
2022 CBC,2025 CBC
2022 CFC,2025 CFC
"NFPA 13, 2022 edition","NFPA 13, 2025 edition"
```

It is written as UTF-8 with a byte-order mark so Excel opens and saves it correctly, and read with or without the mark or the header. Fields are taken exactly as written, so a rule may start or end with a space. A file with the same find text twice, or an empty find text with a replacement, is refused with the row number.

### Change log

Each apply with the change log enabled writes `DocxFindReplace-changelog-<timestamp>.csv` in the scanned folder, one row per match:

| Column | Meaning |
|---|---|
| `file` | Full path of the document |
| `location` | Body, table, header, footer, footnote, endnote or comment, with the paragraph or cell |
| `section` | Headings above the match, for body text |
| `find`, `replace` | The rule |
| `before`, `match`, `after` | The context shown in the preview |
| `status` | `applied`, `skipped: <reason>`, or `not selected` |
| `backup` | The backup taken of that file, if any |

## File Structure

```
docx-find-and-replace/
├── main.py                 # GUI application (tkinter)
├── document_processor.py   # Core scan and replace logic, change log, backups
├── settings.py             # Rule-set CSV files and remembered settings
├── version.py              # Version and product name, used by app and installer
├── requirements.txt        # Runtime dependencies
├── requirements-dev.txt    # Test and packaging dependencies
├── pytest.ini              # Test configuration
├── LICENSE.txt             # Shown by the installer
├── packaging/
│   ├── app.spec            # PyInstaller: one-folder build (wrapped by installer)
│   ├── app-portable.spec   # PyInstaller: single-file portable build
│   ├── installer.iss       # Inno Setup installer script
│   ├── make_version_file.py# Windows file-version resource, from version.py
│   ├── build.bat           # One-command Windows build
│   └── icon.ico            # Application icon
├── .github/workflows/
│   └── build.yml           # Tests on both python-docx versions, then packages
├── tests/
│   ├── conftest.py         # Fixture builders (hyperlinks, merged cells, images, ...)
│   ├── test_traversal.py   # Text extraction and paragraph identity
│   ├── test_replace.py     # Paragraph-level replacement
│   ├── test_search.py      # Occurrence finding, context, folder scanning
│   ├── test_apply.py       # End-to-end scan → apply, backups, reporting
│   ├── test_rules.py       # Rule sets, whole word, overlap conflicts, subfolders
│   ├── test_tracked.py     # Tracked-changes mode, checked by simulated Accept/Reject
│   ├── test_sections.py    # Heading path detection
│   ├── test_changelog.py   # Change log CSV, backup discovery and restore
│   ├── test_settings.py    # Rule-set files and the settings file
│   ├── test_coverage.py    # Notes parts, staleness guard, progress and cancel
│   ├── test_gui.py         # GUI behaviour against the tkinter stub
│   └── tkstub.py           # Headless tkinter stand-in used by test_gui.py
└── README.md               # This file
```

## Document Coverage

Scanned and replaced:

- Body paragraphs
- Tables, including tables nested inside table cells
- Headers and footers (default, first-page, and even-page)
- Footnotes, endnotes and comments
- Text boxes
- Text inside hyperlinks
- Text inside tracked insertions (`w:ins`), smart tags, and content controls
- The cached result of a complex field — a cross-reference or a table of contents entry — which Word stores as ordinary text (see [Known Limitations](#known-limitations): these replacements can revert)

Deliberately not touched:

- Text inside tracked deletions (`w:del`) — this text is already marked for removal
- Word's footnote separator entries — rule lines, not content
- Field codes (`w:instrText`) — editing the instruction would break the field
- Cached results inside `w:fldSimple` — the simple-field form is not descended into

See [Known Limitations](#known-limitations) for content that is not reached at all.

## How It Works

### Rule sets and overlaps

A scan runs every rule over every paragraph and sorts the matches into document order. Each match remembers the rule it came from, so one apply can write a different replacement for each. Two rules can match overlapping text — "2022" and "2022 CBC" both match the start of "2022 CBC" — and which one was meant is not something the tool should guess, so both matches are flagged, cannot be selected, and are refused at apply time even if a match object is selected by hand. The check is pairwise: with "2022 CBC", "2022" and "CBC" all matching, the third overlaps the first even though it does not overlap the second.

Rules are applied to the original text only. A rule whose find text equals another rule's replacement does not chain.

### Tracked-changes mode

In place, a replacement is written straight into the run's text. As a tracked change, the same edit is written the way Word writes one: the old text is wrapped in a `w:del` (its `w:t` becoming `w:delText`) and the new text follows in a `w:ins`, each stamped with the author, a timestamp and a revision id that continues above the highest id already in the document. Word then shows the strike-through and the underline, and Accept All produces exactly what in-place mode would have written; Reject All produces the original.

A run often carries more than text — an inline image, a footnote reference, a field character. Before anything is wrapped, the matched text is split into a run of its own and everything else moves to sibling runs that are never wrapped, so an image sharing a run with the matched text is not marked as deleted. The inserted text takes the formatting of the first run of the matched text, as Word does when typing over a selection.

Two structures are refused rather than guessed at, and reported with a reason: a match inside text that is itself a tracked insertion (Word has no representation for an insertion inside an insertion — accept or reject that revision in Word first, or run without tracked changes), and a match inside a comment (Word does not track revisions in comment text). Running tracked mode twice without accepting in between therefore skips matches inside the first run's insertions.

### Section context

For body text, each paragraph records the headings above it. A paragraph is a heading when its outline level is 0 to 8 — set on the paragraph itself, on its style, or inherited from the style it is based on — with a fallback for styles simply named "heading N". Level 9 is body text and stops inheritance, which is how Word's own TOC Heading style is defined. Table cells and text boxes take the path of the text around them. Headers, footers and notes have no section.

The label is the heading's text as stored. Numbering that Word generates from a list definition — "PART 1", "1.02" — is not stored with the paragraph and so does not appear.

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

### Guarding against a document that changed underneath the preview

A preview records character offsets into one particular version of a file. If
the document is edited in Word between Preview and Apply, those offsets may
still *validate* — the search text can happen to sit at the same offset in
different surrounding prose — and the replacement would land in the wrong place.

Each match therefore records the file's size and modification time at scan time,
and `apply_changes` re-checks both before opening the file. A file that changed
is refused whole, reported as an error, and counted as skipped. No backup is
written for it, because nothing was touched.

### Why the version floor matters

`python-docx` changed the behaviour of `Paragraph.text` between 0.8.x and 1.1: older versions exclude hyperlink text, newer versions include it. Code that depends on that property behaves differently depending on which version `pip` happened to resolve. This tool reads the XML directly rather than depending on that property, but the floor is pinned at 1.1.0 anyway so that the APIs it does use (`part.related_parts`, `section._sectPr`, part element access) are guaranteed present.

## Testing

```bash
pip install -r requirements-dev.txt
pytest
```

The suite builds real `.docx` files on disk for each structure the replacement logic has to get right — hyperlinks, merged cells, headers linked across sections, inline images inside a spanned run, nested tables, text boxes in both VML and `mc:AlternateContent` form, footnotes, endnotes, comments, complex fields, tracked changes, and heading styles — then asserts on the resulting documents. Each test named after a corruption case asserts the document is left *correct*, not merely that the call returned without raising. Tracked-changes output is checked by simulating Word's Accept All and Reject All over the saved file and comparing both against the in-place result and the original.

The GUI is covered too, headlessly. `tests/tkstub.py` is a small tkinter
stand-in, so `test_gui.py` can exercise the window's decisions — which status
message is shown, whether a click toggles a checkbox, whether Apply survives an
error, whether stale results can still be applied — without needing a display.

## Building the Installer

Builds run on Windows. PyInstaller is not a cross-compiler, so a Linux or macOS
machine cannot produce these artifacts.

```bat
pip install -r requirements-dev.txt
packaging\build.bat
```

`build.bat` runs the test suite first and refuses to package a failing build,
then produces:

| Output | What it is |
|---|---|
| `dist\DocxFindReplace\DocxFindReplace.exe` | Installed build (one folder) |
| `dist\DocxFindReplace-Portable.exe` | Single-file portable build |
| `dist\installer\DocxFindReplace-Setup-<version>.exe` | The installer |

The installer step needs [Inno Setup 6](https://jrsoftware.org/isdl.php). If it
is missing, the executables are still built and only the setup step is skipped.

The same build runs in GitHub Actions on every push and pull request, and
attaches the installer and portable exe to the GitHub release when a `v*` tag is
pushed. Version numbers come from `version.py` — bump it there and the
executable's file properties, the installer, and the output filenames all
follow.

### Verifying a build

A packaged build can fail in ways the source never does: python-docx's XML
templates left out of the bundle, or lxml's dynamically resolved imports not
collected. Neither shows up until the document code actually runs.

Both executables therefore support a headless self-check that performs a real
find-and-replace on a temporary document and exits non-zero if anything is
wrong:

```bat
DocxFindReplace.exe --selftest report.txt
```

CI runs this against both built executables before the installer is packaged.

## Known Limitations

- **No regex** — Find text is matched literally (case-sensitive or case-insensitive). **Whole word only** treats letters, digits and the underscore as word characters; a find text whose edge is punctuation, such as "(2022)", needs no boundary on that side
- **Restore is a copy back** — Restore puts the backup's bytes back over the document. Anything done to the document after the backup was taken is lost, and the tool does not write a further backup of the state it is overwriting. Backups are timestamped (`spec.docx.20260909-141530.bak`) and only created for files that are actually modified; renaming one back to `.docx` by hand still works
- **Overlapping rules are flagged, not resolved** — Matches from two rules on overlapping text are never applied, and there is no "longest rule wins" option. A broad rule beside a specific one ("2022" and "2022 CBC") makes every occurrence of the specific one a conflict
- **Rules do not chain** — Replacement text is never re-scanned, so a rule whose find text equals another rule's replacement applies to the original text only
- **Tracked changes** — A match inside an existing tracked insertion, or inside a comment, is skipped with a reason rather than written as a revision. A match that starts outside a hyperlink and ends inside it places its insertion inside the hyperlink. Copied run formatting drops any tracked formatting change (`w:rPrChange`) the original run carried; the original run keeps it. The output is verified by simulating Accept All and Reject All; opening the result in Word is still the final check
- **Section detection needs outline levels** — Templates whose part and article styles carry no outline level and are not named "heading N" show an empty Section column. Word's automatic numbering is not shown in the path
- **Double-click also clicks** — Word's convention: a double-click on a row first delivers the two single clicks, which toggle the checkbox twice and cancel out. A slow double-click can leave the row toggled
- **Text box fallback copies** — Word stores a text box twice, as a modern `mc:Choice` copy and a legacy `mc:Fallback` copy. Only the `mc:Choice` copy is replaced, which is what current versions of Word render; the fallback copy retains the old text until Word next re-saves the file
- **Field results can revert** — Word stores a complex field's cached result (a cross-reference, a table of contents entry) as ordinary text, so it *is* scanned and replaced like any other text. But Word regenerates that text from the field code the next time the field updates, which silently undoes the replacement. For a table of contents, replace the underlying heading text and update the field rather than relying on the cached entry. The field code itself is never touched, and results inside the older `w:fldSimple` form are not scanned at all
- **Preview can go stale** — If a document is edited in Word between clicking Preview and clicking Apply, the file is refused outright rather than edited with offsets that no longer describe it. The completion dialog names the file and reports the skipped changes; re-run the preview to pick up the new content. Detection is by file size and modification time, so an edit that changes neither would go unnoticed.
- **Single-pass matching** — Replacement text is not re-scanned, so replacing "2022" with "2022 and 2025" applies exactly once per original occurrence

## Copyright Notice

**Copyright © 2025 Abraham Borg. All Rights Reserved.**

This software and associated documentation files (the "Software") are the proprietary property of Abraham Borg. 

**Unauthorized copying, modification, distribution, or use of this Software, via any medium, is strictly prohibited without express written permission from the copyright holder.**

This Software is provided for review and reference purposes only. No license or right to use, copy, modify, or distribute this Software for any purpose, commercial or non-commercial, is granted.
