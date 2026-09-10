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
- `version.py` — the one place the version lives. The GUI title, the Windows
  file-version resource, the installer and the output filenames all read it.
- `packaging/` — PyInstaller specs, the Inno Setup script and `build.bat`.
- `.github/workflows/build.yml` — runs the suite against both ends of the
  supported python-docx range, then packages.

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
relationship from the document part (`_NOTE_PARTS`). Any part that is not
registered as an XML part loads as an opaque blob with nothing to walk or edit,
and `_collect_paragraphs` skips it silently — so the module registers all three
as `XmlPart` through `PartFactory.part_type_for` at import.

Register all three explicitly, never only the ones the installed python-docx
happens to miss. python-docx registers a class for comments from 1.2.0 and
never for footnotes or endnotes; the supported range starts at 1.1.0, which
registers none of them. Relying on the library's own registration made comment
coverage depend on which version pip resolved. `setdefault` keeps a
library-supplied class when there is one.

Word's `separator` and `continuationSeparator` footnote entries are rule lines,
not content, and are skipped.

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

**Run the suite against the floor, not just the latest.** The supported range
is a promise, and the versions in it differ in what they register and what
their convenience APIs return. A change that works on the newest python-docx
can silently do nothing on 1.1.0 — that is exactly how comment coverage shipped
broken once already.

```
py -m venv .venv-floor
.venv-floor\Scripts\pip install "python-docx==1.1.0" pytest
.venv-floor\Scripts\python -m pytest
```

`test_note_parts_load_as_xml_on_every_supported_version` asserts the
registration directly, so that particular gap now fails loudly on any version.

## Packaging

Builds are Windows-only; PyInstaller does not cross-compile. There is no way to
produce or test these artifacts from a Linux container, so **the CI workflow is
the verification** — do not hand over build config that has not gone green there.

A frozen build breaks in ways the source never does. python-docx loads XML
templates from its package directory and lxml resolves some imports
dynamically, and PyInstaller sees neither from the import graph:
`collect_data_files("docx")` and the `lxml._elementpath` hidden import in both
specs exist for that reason. Removing either produces a build that starts fine
and then fails the moment a document is opened.

`main.py --selftest [report]` is the guard: it runs a real find-and-replace on a
temporary document, imports tkinter, writes a report and exits non-zero on
failure. The workflow runs it against **both built executables** before
packaging. A windowed executable has no console, hence the report file. If you
change what the packaged build depends on, extend the selftest to cover it —
otherwise CI will keep passing while the shipped exe is broken.

Other things worth not relearning:

- `.gitignore` carries `*.spec` from the standard Python template, which would
  silently swallow `packaging/*.spec`. The `!packaging/*.spec` negation keeps
  the build config tracked; do not remove it.
- Builds are one-folder for the installer and one-file only for the portable
  exe. One-file self-extracts on every launch, which is slower and draws more
  antivirus attention. UPX is off for the same reason.
- The installer is per-user (`PrivilegesRequired=lowest`) so it needs no
  administrator rights on a managed corporate machine. Do not "fix" this to a
  Program Files install.
- Never change `AppId` in `installer.iss`. It is how Windows tells an upgrade
  from a second copy; changing it strands the installed version.
- Bump the version in `version.py` only. Nothing else should hard-code it.
- A release tag must match `version.py`. Everything that names or stamps a
  binary reads `version.py`, but the release is created from the pushed tag,
  and nothing else reconciles them — a `v1.1.0` tag pushed against a 1.0.0
  `version.py` would publish a release labelled 1.1.0 containing binaries
  stamped 1.0.0. `packaging/check_tag.py` fails the build before anything is
  built; it is plain Python precisely so it can be tested here rather than
  trusted because it reads well in YAML.

## Documentation

Update `README.md` when behaviour changes. Its "Known Limitations" section is a
correctness claim about what the tool does *not* reach — keep it truthful and
current, and do not describe coverage the code does not actually have.
