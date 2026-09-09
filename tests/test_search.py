"""Search: occurrence finding, context extraction and folder scanning."""

import os

import pytest
from docx import Document

import document_processor as dp


# ----------------------------------------------------- occurrence finding

def test_occurrences_do_not_overlap():
    """Word counts 'aa' twice in 'aaaa'; advancing by one character counts three."""
    assert dp._find_all_occurrences("aaaa", "aa", True) == [0, 2]


def test_repeated_token_offsets():
    assert dp._find_all_occurrences("2022 2022 2022", "2022", True) == [0, 5, 10]


def test_multi_space_search_does_not_overlap():
    assert dp._find_all_occurrences("a    b", "  ", True) == [1, 3]


def test_case_insensitive_finding():
    assert dp._find_all_occurrences("CBC cbc Cbc", "cbc", False) == [0, 4, 8]
    assert dp._find_all_occurrences("CBC cbc Cbc", "cbc", True) == [4]


def test_empty_needle_finds_nothing():
    assert dp._find_all_occurrences("anything", "", True) == []


# ----------------------------------------------------- context

def test_context_is_not_truncated_twice():
    """
    The ellipsis prefix used to push the string past the display limit, which
    then re-truncated it and silently dropped real characters.
    """
    text = "x" * 200 + "TARGET" + "y" * 200
    before, after = dp._get_context(text, 200, len("TARGET"))
    match = dp.Match(
        file_path="f", location_type="body", location_detail="d",
        paragraph_key="k", char_offset=200,
        context_before=before, match_text="TARGET", context_after=after,
    )
    assert match.display_context == f"{before}[TARGET]{after}"
    assert before.startswith("...")
    assert after.endswith("...")
    assert len(before) <= 50
    assert len(after) <= 50


def test_context_at_paragraph_boundaries():
    before, after = dp._get_context("TARGET tail", 0, len("TARGET"))
    assert before == ""
    assert after == " tail"


def test_display_context_is_single_line():
    """A <w:br/> becomes \\n, which would break a Treeview row."""
    match = dp.Match(
        file_path="f", location_type="body", location_detail="d",
        paragraph_key="k", char_offset=0,
        context_before="line one\n", match_text="2022\tCBC", context_after="\nline two",
    )
    assert "\n" not in match.display_context
    assert "\t" not in match.display_context


# ----------------------------------------------------- folder scanning

def test_scan_finds_matches_across_locations(merged_and_linked_doc):
    results = dp.scan_documents(os.path.dirname(merged_and_linked_doc), "2022 CBC", True)
    assert len(results) == 1
    matches = results[0].matches
    # body + merged cell + plain cell + shared header, each exactly once
    assert len(matches) == 4
    assert {m.location_type for m in matches} == {'body', 'table', 'header'}


def test_scan_does_not_inflate_counts(merged_and_linked_doc):
    """Preview count must equal the number of physical occurrences."""
    results = dp.scan_documents(os.path.dirname(merged_and_linked_doc), "2022", True)
    assert results[0].match_count == 4


def test_scan_reports_files_scanned_separately_from_matches(repeated_doc):
    """
    'No .docx files in this folder' and 'no matches in the files here' are
    different answers; the caller needs to be able to tell them apart.
    """
    folder = os.path.dirname(repeated_doc)
    results, summary = dp.scan_documents_detailed(folder, "NOTHING_MATCHES_THIS", True)
    assert results == []
    assert summary['files_scanned'] == 1
    assert summary['files_with_matches'] == 0


def test_scan_empty_folder(docdir):
    results, summary = dp.scan_documents_detailed(str(docdir), "anything", True)
    assert results == []
    assert summary['files_scanned'] == 0


def test_scan_skips_word_lock_files(docdir, repeated_doc):
    lock = docdir / "~$repeated.docx"
    lock.write_bytes(b"not a real docx")
    results, summary = dp.scan_documents_detailed(str(docdir), "2022", True)
    assert summary['files_scanned'] == 1
    assert all(not r.file_name.startswith("~$") for r in results)


def test_scan_reports_corrupt_file_as_error(docdir):
    bad = docdir / "broken.docx"
    bad.write_bytes(b"this is not a zip archive")
    results, summary = dp.scan_documents_detailed(str(docdir), "2022", True)
    assert len(results) == 1
    assert results[0].error
    assert results[0].matches == []


def test_scan_with_empty_find_text_returns_nothing(repeated_doc):
    results, summary = dp.scan_documents_detailed(os.path.dirname(repeated_doc), "", True)
    assert results == []


def test_match_text_records_actual_casing(docdir):
    doc = Document()
    doc.add_paragraph("The CBC and the cbc")
    path = docdir / "case.docx"
    doc.save(str(path))

    results = dp.scan_documents(str(docdir), "cbc", False)
    assert [m.match_text for m in results[0].matches] == ["CBC", "cbc"]
