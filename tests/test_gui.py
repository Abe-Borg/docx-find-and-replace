"""
GUI behaviour, exercised headlessly against the tkinter stub in `tkstub.py`.

These cover the decisions the window makes rather than how it looks: which
status message is shown, whether a click toggles a checkbox, whether Apply
survives an error, and whether stale results can still be applied.
"""

import tkstub

tkstub.install()          # must precede `import main`

import pytest             # noqa: E402
import main               # noqa: E402
import document_processor as dp  # noqa: E402


def make_match(location_type='body', detail='Paragraph 1', selected=True,
               file_path='C:\\specs\\a.docx'):
    return dp.Match(
        file_path=file_path,
        location_type=location_type,
        location_detail=detail,
        paragraph_key='/word/document.xml#0',
        char_offset=0,
        context_before='Per the ',
        match_text='2022 CBC',
        context_after=' requirements.',
        is_selected=selected,
    )


@pytest.fixture
def app():
    tkstub.dialogs.reset()
    root = main.tk.Tk()
    return main.FindReplaceApp(root)


@pytest.fixture
def previewed(app):
    """An app showing one file with three matches, as after a preview."""
    fr = dp.FileResult(file_path='C:\\specs\\a.docx', file_name='a.docx',
                       matches=[make_match(), make_match(), make_match()])
    app.scan_find_text = '2022 CBC'
    app.scan_case_sensitive = True
    app.find_var._value = '2022 CBC'
    app._display_results([fr], {'files_scanned': 1, 'files_with_matches': 1,
                                'cancelled': False})
    return app


# ----------------------------------------------------- labels

@pytest.mark.parametrize("location_type,expected", [
    ('body', ''),
    ('table', '[Table] '),
    ('header', '[Header] '),
    ('footer', '[Footer] '),
    ('footnote', '[Footnote] '),
    ('endnote', '[Endnote] '),
    ('comment', '[Comment] '),
])
def test_every_location_type_has_a_label(app, location_type, expected):
    """A location the tree cannot name would show up as a bare, ambiguous row."""
    label = app._match_label(make_match(location_type=location_type))
    assert label.startswith(f"{main.CHECKED} {expected}")


def test_label_reflects_checkbox_state(app):
    assert app._match_label(make_match(selected=True)).startswith(main.CHECKED)
    assert app._match_label(make_match(selected=False)).startswith(main.UNCHECKED)


# ----------------------------------------------------- status messages

def test_no_docx_files_message(app):
    app._display_results([], {'files_scanned': 0, 'cancelled': False})
    assert app.progress_var.get() == "No .docx files found in the selected folder"


def test_no_matches_in_a_folder_full_of_files(app):
    """The old message claimed the folder was empty. It was not."""
    app._display_results([], {'files_scanned': 200, 'cancelled': False})
    assert "No matches found in 200 .docx files" == app.progress_var.get()


def test_cancelled_scan_says_so(app):
    app._display_results([], {'files_scanned': 12, 'cancelled': True})
    assert "cancelled after 12 files" in app.progress_var.get()


def test_cancelled_scan_with_partial_results(previewed):
    fr = dp.FileResult(file_path='b.docx', file_name='b.docx', matches=[make_match()])
    previewed._display_results([fr], {'files_scanned': 3, 'cancelled': True})
    status = previewed.progress_var.get()
    assert "Found 1 match" in status
    assert "cancelled after 3 files" in status


# ----------------------------------------------------- tree interaction

def test_clicking_the_expander_does_not_toggle_the_file(previewed):
    """
    identify_row alone ignores x, so a click on the expand arrow used to flip
    every checkbox in the file.
    """
    file_node = previewed.tree.get_children()[0]
    before = [m.is_selected for m in previewed.all_matches]

    previewed.tree.next_identify_element = "Treeitem.indicator"
    previewed.tree.next_identify_row = file_node
    previewed._on_tree_click(type("E", (), {"x": 12, "y": 10})())

    assert [m.is_selected for m in previewed.all_matches] == before


def test_clicking_a_file_row_toggles_all_its_matches(previewed):
    file_node = previewed.tree.get_children()[0]
    previewed.tree.next_identify_element = "text"
    previewed.tree.next_identify_row = file_node
    previewed._on_tree_click(type("E", (), {"x": 80, "y": 10})())

    assert all(not m.is_selected for m in previewed.all_matches)


def test_clicking_a_match_toggles_only_that_match(previewed):
    file_node = previewed.tree.get_children()[0]
    first_child = previewed.tree.get_children(file_node)[0]
    previewed.tree.next_identify_element = "text"
    previewed.tree.next_identify_row = first_child
    previewed._on_tree_click(type("E", (), {"x": 80, "y": 30})())

    assert [m.is_selected for m in previewed.all_matches] == [False, True, True]


def test_tree_is_inert_while_applying(previewed):
    """The worker thread reads the same Match objects the tree would mutate."""
    file_node = previewed.tree.get_children()[0]
    previewed.processing_kind = 'apply'
    previewed.tree.next_identify_element = "text"
    previewed.tree.next_identify_row = file_node

    result = previewed._on_tree_click(type("E", (), {"x": 80, "y": 10})())

    assert result == "break"
    assert all(m.is_selected for m in previewed.all_matches)


def test_select_all_is_inert_while_applying(previewed):
    previewed._deselect_all()
    previewed.processing_kind = 'apply'
    previewed._select_all()
    assert all(not m.is_selected for m in previewed.all_matches)


# ----------------------------------------------------- stale settings

def test_editing_the_search_term_invalidates_the_preview(previewed):
    """
    Offsets belong to the term they were found with. Leaving them applicable
    under a new term would validate old offsets against new rules.
    """
    previewed.find_var.set("2025 CBC")

    assert previewed.all_matches == []
    assert previewed.tree.get_children() == ()
    assert previewed.apply_btn.state == "disabled"
    assert "run Preview Changes again" in previewed.progress_var.get()


def test_toggling_case_sensitivity_invalidates_the_preview(previewed):
    previewed.case_var.set(False)
    assert previewed.all_matches == []
    assert "run Preview Changes again" in previewed.progress_var.get()


def test_changing_the_replacement_text_keeps_the_preview(previewed):
    """Replacement text is not used until apply, so it does not invalidate."""
    previewed.replace_var.set("2025 CBC")
    assert len(previewed.all_matches) == 3
    assert previewed.apply_btn.state == "normal"


def test_confirm_dialog_shows_the_term_that_was_scanned(previewed, monkeypatch):
    """
    The dialog used to print the live entry box, so it could name a search that
    was never run.
    """
    monkeypatch.setattr(main.threading, "Thread", lambda **kwargs: type(
        "T", (), {"start": lambda self: None})())
    previewed.find_var._value = "something else entirely"   # bypass the trace

    previewed._start_apply()

    _title, message = tkstub.dialogs.questions[-1]
    assert 'Find: "2022 CBC"' in message
    assert "something else entirely" not in message
    assert "Case sensitive: yes" in message


def test_apply_uses_the_case_setting_from_the_scan(previewed, monkeypatch):
    captured = {}

    def fake_apply(matches, replace_text, **kwargs):
        captured.update(kwargs)
        return {'total_replaced': 0, 'total_skipped': 0, 'files_modified': 0,
                'backups': [], 'errors': []}

    monkeypatch.setattr(main, "apply_changes", fake_apply)
    previewed.scan_case_sensitive = True
    previewed.case_var._value = False        # changed after the scan

    previewed._run_apply()

    assert captured['case_sensitive'] is True


# ----------------------------------------------------- apply reporting

def test_skipped_changes_are_reported(previewed):
    previewed._apply_complete({'total_replaced': 6, 'total_skipped': 3,
                               'files_modified': 2, 'backups': ['a', 'b'],
                               'errors': []})
    _title, message = tkstub.dialogs.info[-1]
    assert "3 selected changes were NOT applied" in message
    assert "Backups created: 2" in message
    assert "(3 skipped)" in previewed.progress_var.get()


def test_clean_run_shows_no_warning(previewed):
    previewed._apply_complete({'total_replaced': 4, 'total_skipped': 0,
                               'files_modified': 1, 'backups': ['a'],
                               'errors': []})
    _title, message = tkstub.dialogs.info[-1]
    assert "NOT applied" not in message


def test_errors_are_listed(previewed):
    previewed._apply_complete({'total_replaced': 0, 'total_skipped': 3,
                               'files_modified': 0, 'backups': [],
                               'errors': ["a.docx: was modified on disk since the preview."]})
    _title, message = tkstub.dialogs.info[-1]
    assert "was modified on disk" in message


def test_results_are_cleared_after_a_successful_apply(previewed):
    previewed._apply_complete({'total_replaced': 3, 'total_skipped': 0,
                               'files_modified': 1, 'backups': [], 'errors': []})
    assert previewed.all_matches == []
    assert previewed.apply_btn.state == "disabled"


def test_apply_error_leaves_apply_available_for_retry(previewed):
    """Losing the button meant re-scanning the whole folder to try again."""
    previewed.processing_kind = 'apply'
    previewed._apply_error("disk full")

    assert previewed.apply_btn.state == "normal"
    assert len(previewed.all_matches) == 3
    assert tkstub.dialogs.errors


# ----------------------------------------------------- shutdown

def test_closing_during_apply_is_refused(previewed):
    previewed.processing_kind = 'apply'
    previewed._on_close()

    assert previewed.root.destroyed is False
    assert tkstub.dialogs.warnings


def test_closing_during_scan_cancels_then_quits(app):
    import threading
    app.processing_kind = 'scan'
    app.scan_cancel = threading.Event()
    tkstub.dialogs.answer = True

    app._on_close()

    assert app.scan_cancel.is_set()
    assert app.close_when_idle is True
    assert app.root.destroyed is False      # not until the scan actually stops

    app._finish_processing()
    assert app.root.destroyed is True


def test_declining_the_scan_close_prompt_keeps_running(app):
    import threading
    app.processing_kind = 'scan'
    app.scan_cancel = threading.Event()
    tkstub.dialogs.answer = False

    app._on_close()

    assert not app.scan_cancel.is_set()
    assert app.close_when_idle is False
    assert app.root.destroyed is False


def test_closing_while_idle_quits(app):
    app._on_close()
    assert app.root.destroyed is True


def test_results_survive_a_scan_that_completes_after_a_close_request(app):
    """
    _finish_processing may destroy the window, so it has to run last - otherwise
    the tree is populated after the root is gone.
    """
    import threading
    app.processing_kind = 'scan'
    app.scan_cancel = threading.Event()
    app.close_when_idle = True

    fr = dp.FileResult(file_path='a.docx', file_name='a.docx', matches=[make_match()])
    app._display_results([fr], {'files_scanned': 1, 'cancelled': True})

    assert app.root.destroyed is True
    assert len(app.all_matches) == 1        # populated before the window went away
    assert "cancelled after 1 file" in app.progress_var.get()
