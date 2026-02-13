# Word Document Batch Find & Replace

A Python GUI application for performing batch find-and-replace operations across multiple Word documents (.docx) with a preview-before-commit workflow.

## Use Case

Updating technical specifications and building code references across multiple Word documents — for example, replacing "2022 CBC" with "2025 CBC" while being able to skip historical references like "designed in 2022."

## Features

- **Preview before commit** — Scan all documents and review every match with surrounding context before applying any changes
- **Selective replacement** — Check/uncheck individual matches or entire files; only approved changes are applied
- **Cross-run matching** — Finds text even when Word has split it across multiple formatting runs (a common issue with `python-docx`)
- **Full document coverage** — Scans body paragraphs, tables, headers, and footers
- **Formatting preservation** — Replacements preserve bold, italic, font, color, and other run-level formatting
- **Backup creation** — Optional `.docx.bak` backup files before modifying originals
- **Case sensitivity toggle** — Search with or without case sensitivity
- **Error handling** — Gracefully skips corrupted, locked, or permission-denied files with clear error messages

## Requirements

- Python 3.8+
- Windows (tested), macOS/Linux (should work with tkinter installed)

## Installation

```bash
cd word-find-replace
pip install -r requirements.txt
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
6. **Apply** — Click "Apply Selected Changes." A confirmation dialog shows the summary before proceeding.

## File Structure

```
word-find-replace/
├── main.py                 # GUI application (tkinter)
├── document_processor.py   # Core scan and replace logic
├── requirements.txt        # Python dependencies
└── README.md               # This file
```

## How Cross-Run Matching Works

Word internally stores paragraph text as a sequence of "runs" — segments that share the same formatting. A single phrase like "2022 CBC" can be split across multiple runs due to spell-check history, partial edits, or even just opening and re-saving. For example:

```
Run 1: "Reference the 20"  (bold)
Run 2: "22 CBC"            (bold + red)
Run 3: " for compliance."  (normal)
```

A naive `run.text.replace()` approach would miss this entirely. This tool instead:

1. Concatenates all run texts to get the full paragraph text for searching
2. Builds a character-to-run mapping to locate which runs contain each match
3. Performs replacement across the correct runs while preserving each run's formatting

## Known Limitations

- **Non-recursive** — Only processes `.docx` files in the selected folder, not subfolders
- **No regex** — Find text is matched literally (case-sensitive or case-insensitive)
- **No undo** — Use the backup feature; there is no built-in undo
- **Cross-run edge case** — If the find text is split across runs in an unusual way (e.g., a run boundary falls in the middle of a character in a multi-byte encoding), replacement may not work correctly. This is extremely rare in practice.
