"""End-to-end: scan then apply, including backup and reporting behaviour."""

import glob
import os

import pytest
from docx import Document

import document_processor as dp
from conftest import all_text, count_drawings


def _scan(path, find, case=True):
    results = dp.scan_documents(os.path.dirname(path), find, case)
    return [m for r in results for m in r.matches]


# ----------------------------------------------------- correctness

def test_hyperlink_paragraph_is_not_corrupted(hyperlink_doc):
    matches = _scan(hyperlink_doc, "2022 CBC")
    result = dp.apply_changes(matches, "2025 CBC", create_backups=False)
    assert result['total_replaced'] == 1
    assert result['errors'] == []
    assert all_text(hyperlink_doc)[0] == "See 2025 CBC link and done."


def test_merged_cell_and_linked_header_replaced_exactly_once(merged_and_linked_doc):
    matches = _scan(merged_and_linked_doc, "2022")
    assert len(matches) == 4
    result = dp.apply_changes(matches, "2022 and 2025", create_backups=False)
    assert result['total_replaced'] == 4

    texts = all_text(merged_and_linked_doc)
    merged = next(t for t in texts if t.startswith("Merged:"))
    header = next(t for t in texts if t.startswith("HEADER"))
    assert merged == "Merged: 2022 and 2025 CBC note"
    assert header == "HEADER 2022 and 2025 CBC"


def test_inline_image_survives_apply(image_in_run_doc):
    matches = _scan(image_in_run_doc, "2022 CBC")
    dp.apply_changes(matches, "2025 CBC", create_backups=False)
    assert count_drawings(image_in_run_doc) == 1
    assert all_text(image_in_run_doc)[0] == "See Figure 2025 CBC ref"


def test_nested_table_is_replaced(nested_table_doc):
    matches = _scan(nested_table_doc, "2022 CBC")
    result = dp.apply_changes(matches, "2025 CBC", create_backups=False)
    assert result['total_replaced'] == 1
    assert any("Nested 2025 CBC" in t for t in all_text(nested_table_doc))


def test_textbox_is_replaced(textbox_doc):
    matches = _scan(textbox_doc, "2022 CBC")
    result = dp.apply_changes(matches, "2025 CBC", create_backups=False)
    assert result['total_replaced'] == 1
    assert any("Textbox 2025 CBC" in t for t in all_text(textbox_doc))


def test_deselected_matches_are_left_alone(repeated_doc):
    matches = _scan(repeated_doc, "2022")
    assert len(matches) == 3
    matches[0].is_selected = False
    matches[2].is_selected = False

    result = dp.apply_changes(matches, "2025", create_backups=False)
    assert result['total_replaced'] == 1
    assert all_text(repeated_doc)[0] == "aaaa and 2022 2025 2022"


def test_nothing_selected_changes_nothing(repeated_doc):
    matches = _scan(repeated_doc, "2022")
    for m in matches:
        m.is_selected = False
    before = all_text(repeated_doc)
    result = dp.apply_changes(matches, "2025", create_backups=True)
    assert result == {'total_replaced': 0, 'total_skipped': 0, 'files_modified': 0,
                      'backups': [], 'errors': []}
    assert all_text(repeated_doc) == before


def test_offsets_stay_valid_across_the_reload(merged_and_linked_doc):
    """
    Scan and apply open the file separately. If paragraph identity did not
    survive the reload, replacements would land in the wrong paragraph.
    """
    matches = _scan(merged_and_linked_doc, "2022 CBC")
    keys = {m.paragraph_key for m in matches}
    assert len(keys) == 4
    result = dp.apply_changes(matches, "2025 CBC", create_backups=False)
    assert result['total_replaced'] == 4
    assert result['total_skipped'] == 0


# ----------------------------------------------------- reporting

def test_every_selected_match_is_accounted_for(repeated_doc):
    matches = _scan(repeated_doc, "2022")
    result = dp.apply_changes(matches, "2025", create_backups=False)
    assert result['total_replaced'] + result['total_skipped'] == len(matches)
    assert all(m.applied is True for m in matches)


def test_skipped_matches_are_reported_not_swallowed(repeated_doc):
    """
    A selected match that cannot be applied has to show up in the result, or the
    dialog reports fewer replacements than requested with no explanation.
    """
    matches = _scan(repeated_doc, "2022")
    matches[1].char_offset = 1  # no longer points at "2022"

    result = dp.apply_changes(matches, "2025", create_backups=False)
    assert result['total_replaced'] == 2
    assert result['total_skipped'] == 1
    assert [m.applied for m in matches] == [True, False, True]


def test_missing_paragraph_is_skipped_not_misapplied(repeated_doc):
    matches = _scan(repeated_doc, "2022")
    matches[0].paragraph_key = "/word/document.xml#9999"
    result = dp.apply_changes(matches, "2025", create_backups=False)
    assert result['total_replaced'] == 2
    assert result['total_skipped'] == 1


def test_unreadable_file_is_reported(docdir):
    bad = docdir / "broken.docx"
    bad.write_bytes(b"not a zip")
    match = dp.Match(
        file_path=str(bad), location_type="body", location_detail="Paragraph 1",
        paragraph_key="/word/document.xml#0", char_offset=0,
        context_before="", match_text="2022", context_after="",
    )
    result = dp.apply_changes([match], "2025", create_backups=False)
    assert result['total_replaced'] == 0
    assert result['total_skipped'] == 1
    assert result['errors']
    assert match.applied is False


# ----------------------------------------------------- backups

def test_backup_is_timestamped_and_never_overwritten(repeated_doc):
    """
    A fixed '.bak' name means the second run overwrites the first run's backup
    with already-modified content, destroying the only copy of the original.
    """
    original = open(repeated_doc, 'rb').read()

    dp.apply_changes(_scan(repeated_doc, "2022"), "2023", create_backups=True)
    dp.apply_changes(_scan(repeated_doc, "2023"), "2024", create_backups=True)

    backups = sorted(glob.glob(repeated_doc + ".*.bak"))
    assert len(backups) == 2
    assert any(open(b, 'rb').read() == original for b in backups), \
        "the true original must still be recoverable after a second run"


def test_backup_paths_are_returned(repeated_doc):
    result = dp.apply_changes(_scan(repeated_doc, "2022"), "2025", create_backups=True)
    assert len(result['backups']) == 1
    assert os.path.exists(result['backups'][0])


def test_no_backup_when_nothing_is_replaced(repeated_doc, docdir):
    """A file whose selected matches all fail should not leave a stray .bak."""
    matches = _scan(repeated_doc, "2022")
    for m in matches:
        m.char_offset = 1  # every offset now stale

    result = dp.apply_changes(matches, "2025", create_backups=True)
    assert result['total_replaced'] == 0
    assert glob.glob(repeated_doc + ".*.bak") == []


def test_backup_disabled_creates_no_files(repeated_doc):
    dp.apply_changes(_scan(repeated_doc, "2022"), "2025", create_backups=False)
    assert glob.glob(repeated_doc + ".*.bak") == []


def test_document_still_opens_after_apply(merged_and_linked_doc):
    dp.apply_changes(_scan(merged_and_linked_doc, "2022"), "2025", create_backups=False)
    Document(merged_and_linked_doc)  # raises if the saved package is invalid


def test_no_temp_files_left_behind(repeated_doc, docdir):
    dp.apply_changes(_scan(repeated_doc, "2022"), "2025", create_backups=False)
    leftovers = [f for f in os.listdir(str(docdir)) if f.endswith(".tmp")]
    assert leftovers == []


def test_progress_callback_is_invoked(repeated_doc):
    seen = []
    dp.apply_changes(_scan(repeated_doc, "2022"), "2025",
                     create_backups=False,
                     progress_callback=lambda name, i, total: seen.append((name, i, total)))
    assert seen
    assert seen[-1][0] == "Done"
