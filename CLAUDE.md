# CLAUDE.md

Guidance for Claude Code when working in this repository.

## What this is

A tkinter desktop tool that performs batch find-and-replace across `.docx` files
in a folder, with a preview-before-commit workflow. It **modifies the user's
documents in place**. Those documents are fire sprinkler and building code
specification sets for nonresidential projects, where a silent wrong edit is
worse than a loud failure.

- `main.py` — tkinter GUI. Scanning and applying run on daemon worker threads;
  all widget updates go back to the main thread via `root.after`.
- `document_processor.py` — all document logic. No tkinter imports; it is
  independently testable and that is deliberate.
- `tests/` — pytest suite that builds real `.docx` fixtures on disk, plus
  `tkstub.py`, a headless tkinter stand-in that lets `test_gui.py` exercise the
  window's decisions without a display.

Target platform is Windows. Paths and any shell examples should assume Windows
even though the code is platform-neutral.

## The one architectural rule

**Paragraph text, character offsets, and the XML nodes holding the text must all
come from a single traversal.** `_paragraph_text_and_nodes()` is that traversal.

Every silent-corruption bug this codebase has had came from mixing two views of
a document that were assumed to agree and did not:

- `Paragraph.text` is built from `w:r | w:hyperlink`; `Paragraph.runs` returns
  only direct `w:r` children. Searching one and editing the other misplaces
  every replacement in a paragraph containing a hyperlink.
- `row.cells` yields the same merged cell once per spanned grid column.
- `section.header` returns the same inherited header for every section linked to
  the previous one.

Do not reintroduce `paragraph.text`, `paragraph.runs`, `doc.paragraphs`,
`doc.tables` or `row.cells` into the scan or replace paths. Walk the XML.

## Rules that exist because breaking them destroyed documents

1. **Never assign to `Run.text`.** The setter calls `clear_content()`, which
   strips every child of the run — inline images, footnote references, comment
   anchors, field characters. Write to `w:t` element text via `_set_node_text`,
   which also sets `xml:space="preserve"`.
2. **Never key a paragraph by list position or `id()`.** Scan and apply open the
   file separately, so identity must be structural: `"<part name>#<ordinal>"`
   from `_collect_paragraphs`. `id(para._p)` is fine for deduping inside one
   traversal and useless across loads.
3. **Never reach headers through `section.header`.** That property *creates* an
   empty header part as a side effect. Resolve `w:headerReference` /
   `w:footerReference` through `doc.part.related_parts` instead — which also
   dedupes linked headers for free, since a linked section carries no reference
   of its own.
4. **Never back up to a fixed filename.** `file + '.bak'` means the second run
   overwrites the first run's backup with already-modified content and the true
   original becomes unrecoverable. Use `_make_backup_path` (timestamped, and it
   probes for collisions).
5. **Back up only when the file is actually about to change,** and save through
   `_save_atomically` (temp file in the same directory, then `os.replace`).
   `python-docx` streams the zip over the destination path, so an interrupted
   direct save leaves a truncated, unopenable document.
6. **Never let a selected change vanish quietly.** `_replace_in_paragraph`
   returns `(applied, skipped)`, `apply_changes` sets `Match.applied` and reports
   `total_skipped`, and the GUI warns when it is non-zero. A validation failure
   must skip, never guess.
7. **Re-verify before editing.** `_replace_in_paragraph` re-reads the paragraph
   and confirms the search text is still at the recorded offset before touching
   anything. Offsets are applied in descending order so earlier ones stay valid.
8. **Refuse a file that changed since the scan.** Each Match records the file's
   size and mtime; `_describe_staleness` re-checks both before the file is
   opened for writing. A stale offset can still validate against different text,
   so the whole file is refused rather than partially edited. Matches built by
   hand with no recorded stat skip the check.
9. **Never read a setting live at apply time.** The GUI stores `scan_find_text`
   and `scan_case_sensitive` when a scan starts and uses those when applying.
   Reading the widgets instead would validate offsets found under one setting
   against another. Changing the folder, search text or case setting discards
   the on-screen results for the same reason.
10. **The results tree is inert while a worker runs.** `_on_tree_click`,
    `_select_all` and `_deselect_all` return early when `is_processing`, because
    the worker thread reads the same Match objects those handlers mutate.

## Traversal specifics

`_INLINE_CONTAINERS` is an **allowlist** of elements to descend into when
gathering a paragraph's text. This is intentional. A denylist would eventually
walk into `w:drawing`, `w:pict` or `mc:AlternateContent` and pull text-box
content into the anchoring paragraph's text, corrupting every offset in it.
Add to the allowlist only after confirming the element cannot contain a nested
`w:txbxContent`.

Text boxes are collected as their own paragraphs by `_walk_textboxes`, which
skips `mc:Fallback` copies (Word writes each text box twice) and skips nested
text boxes that the inner walk will reach on its own.

`w:del` and `w:moveFrom` are excluded: that text is already marked deleted and
must not be edited. `w:fldSimple` is excluded, so its cached result is not
scanned.

Footnotes, endnotes and comments live in their own parts, reached by
relationship from the document part (`_NOTE_PARTS`). python-docx registers a
class for comments but not for footnotes or endnotes, so those would load as
opaque blobs with nothing to walk or edit; the module registers them as
`XmlPart` through `PartFactory.part_type_for` at import. That is the documented
extension point — use `setdefault` so a future python-docx that ships its own
class keeps precedence. Word's `separator` and `continuationSeparator` footnote
entries are rule lines, not content, and are skipped.

Note the asymmetry: a **complex** field (`w:fldChar` begin / `w:instrText` /
separate / result / end) keeps its cached result in ordinary `w:r`/`w:t` runs,
so the walk does scan and replace it. Word regenerates that text on field
update, which reverts the replacement — README documents this under "Field
results can revert". Only `w:instrText` (the field code) is never read. If you
change either behaviour, change the README bullet in the same commit.

## Testing

```
pip install -r requirements-dev.txt
pytest
```

Every fix for a corruption bug needs a test that builds a real `.docx`
exercising the structure and asserts the **resulting document is correct** —
not merely that the call returned without raising. All four Tier 1 bugs returned
`replaced=1, errors=[]` while destroying content, so a green return value proves
nothing. `tests/conftest.py` has builders for hyperlinks, merged cells, linked
headers, images inside a spanned run, nested tables, text boxes (VML and
`mc:AlternateContent`), footnotes/endnotes/comments, complex fields and tracked
changes; extend it rather than hand-rolling XML in a test.

For GUI work, `tests/test_gui.py` installs `tkstub` **before** importing `main`,
so the stub must be in place first. The stub is deliberately minimal: a widget
method `main.py` starts calling that the stub lacks fails loudly with
AttributeError, which is the signal to add it rather than to weaken the test.

## Dependencies

`requirements.txt` pins `python-docx>=1.1.0,<2.0`. Keep a floor and a ceiling:
`>=0.8.11` previously spanned a behavioural change in `Paragraph.text` and made
the failure mode depend on whatever `pip` resolved that day. Update
`requirements.txt` whenever runtime dependencies change and
`requirements-dev.txt` for test dependencies.

## Documentation

Update `README.md` when behaviour changes. Its "Known Limitations" section is a
correctness claim about what the tool does *not* reach — keep it truthful and
current, and do not describe coverage the code does not actually have.
