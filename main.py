"""
main.py
Word Document Batch Find & Replace - GUI Application

A tkinter-based GUI for performing batch find-and-replace operations
across multiple Word documents (.docx) with a preview-before-commit workflow.

Threading model: scanning and applying run on daemon worker threads; every
widget update is marshalled back to the main thread with `root.after`. The
results tree is inert while a worker is running, because the worker reads the
same Match objects the tree would mutate.
"""

import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from document_processor import scan_documents_detailed, apply_changes, FileResult, Match
from typing import List, Optional


# Prefix shown before each match's location in the results tree.
LOCATION_PREFIXES = {
    'table': '[Table] ',
    'header': '[Header] ',
    'footer': '[Footer] ',
    'footnote': '[Footnote] ',
    'endnote': '[Endnote] ',
    'comment': '[Comment] ',
}

CHECKED = "☑"      # ☑
UNCHECKED = "☐"    # ☐
PARTIAL = "☒"      # ☒
WARNING = "⚠"      # ⚠


class FindReplaceApp:
    """Main application window."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Word Document Find & Replace")
        self.root.geometry("900x700")
        self.root.minsize(700, 500)

        # State
        self.scan_results: List[FileResult] = []
        self.all_matches: List[Match] = []
        self.match_tree_map = {}   # tree item id -> Match object
        self.file_tree_map = {}    # tree item id -> FileResult object

        # 'scan', 'apply', or None. Guards the tree and the close button.
        self.processing_kind: Optional[str] = None
        self.scan_cancel: Optional[threading.Event] = None
        self.close_when_idle = False

        # The settings the current results were produced with. Apply uses these
        # rather than reading the widgets live, so a setting changed after the
        # preview cannot be applied to offsets found under the old one.
        self.scan_find_text = ""
        self.scan_case_sensitive = True

        # Configure styles
        style = ttk.Style()
        style.configure("TButton", padding=4)
        style.configure("Header.TLabel", font=("Segoe UI", 10, "bold"))
        style.configure("Status.TLabel", font=("Segoe UI", 9))

        self._build_ui()

        # Changing what gets searched invalidates the results that are on screen.
        for var in (self.folder_var, self.find_var, self.case_var):
            var.trace_add("write", self._on_search_settings_changed)

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    @property
    def is_processing(self) -> bool:
        return self.processing_kind is not None

    def _build_ui(self):
        """Build the complete GUI layout."""
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Input Section ---
        input_frame = ttk.LabelFrame(main_frame, text="Search Settings", padding=8)
        input_frame.pack(fill=tk.X, pady=(0, 8))

        # Folder row
        folder_frame = ttk.Frame(input_frame)
        folder_frame.pack(fill=tk.X, pady=2)
        ttk.Label(folder_frame, text="Folder:", width=8).pack(side=tk.LEFT)
        self.folder_var = tk.StringVar()
        self.folder_entry = ttk.Entry(folder_frame, textvariable=self.folder_var)
        self.folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        self.browse_btn = ttk.Button(folder_frame, text="Browse...", command=self._browse_folder)
        self.browse_btn.pack(side=tk.LEFT)

        # Find row
        find_frame = ttk.Frame(input_frame)
        find_frame.pack(fill=tk.X, pady=2)
        ttk.Label(find_frame, text="Find:", width=8).pack(side=tk.LEFT)
        self.find_var = tk.StringVar()
        self.find_entry = ttk.Entry(find_frame, textvariable=self.find_var)
        self.find_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Replace row
        replace_frame = ttk.Frame(input_frame)
        replace_frame.pack(fill=tk.X, pady=2)
        ttk.Label(replace_frame, text="Replace:", width=8).pack(side=tk.LEFT)
        self.replace_var = tk.StringVar()
        self.replace_entry = ttk.Entry(replace_frame, textvariable=self.replace_var)
        self.replace_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Options and preview row
        options_frame = ttk.Frame(input_frame)
        options_frame.pack(fill=tk.X, pady=(6, 2))

        self.backup_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Create backups (timestamped .bak)",
                        variable=self.backup_var).pack(side=tk.LEFT)

        self.case_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Case sensitive",
                        variable=self.case_var).pack(side=tk.LEFT, padx=(16, 0))

        self.preview_btn = ttk.Button(options_frame, text="Preview Changes",
                                       command=self._start_preview)
        self.preview_btn.pack(side=tk.RIGHT)

        # --- Results Section ---
        results_frame = ttk.LabelFrame(main_frame, text="Results", padding=8)
        results_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        # Toolbar for select all / deselect all
        toolbar = ttk.Frame(results_frame)
        toolbar.pack(fill=tk.X, pady=(0, 4))
        self.select_all_btn = ttk.Button(toolbar, text="Select All", command=self._select_all)
        self.select_all_btn.pack(side=tk.LEFT, padx=(0, 4))
        self.deselect_all_btn = ttk.Button(toolbar, text="Deselect All", command=self._deselect_all)
        self.deselect_all_btn.pack(side=tk.LEFT)
        self.match_count_label = ttk.Label(toolbar, text="", style="Status.TLabel")
        self.match_count_label.pack(side=tk.RIGHT)

        # Treeview with checkboxes
        tree_frame = ttk.Frame(results_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(tree_frame, show="tree", selectmode="browse")
        self.tree.column("#0", width=350, minwidth=200)

        # Scrollbars
        v_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scroll = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        # Click handler for toggling checkboxes
        self.tree.bind("<Button-1>", self._on_tree_click)

        # --- Bottom Section ---
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X)

        self.apply_btn = ttk.Button(bottom_frame, text="Apply Selected Changes",
                                     command=self._start_apply, state=tk.DISABLED)
        self.apply_btn.pack(side=tk.RIGHT)

        self.progress_var = tk.StringVar(value="Ready")
        self.status_label = ttk.Label(bottom_frame, textvariable=self.progress_var,
                                       style="Status.TLabel")
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

    # ------------------------------------------------------------------
    # Input handling
    # ------------------------------------------------------------------

    def _browse_folder(self):
        """Open folder selection dialog."""
        folder = filedialog.askdirectory(title="Select folder containing .docx files")
        if folder:
            self.folder_var.set(folder)

    def _validate_inputs(self) -> bool:
        """Check that folder and find text are provided."""
        folder = self.folder_var.get().strip()
        find_text = self.find_var.get()

        if not folder:
            messagebox.showwarning("Missing Input", "Please select a folder.")
            return False
        if not os.path.isdir(folder):
            messagebox.showwarning("Invalid Folder", "The selected folder does not exist.")
            return False
        if not find_text:
            messagebox.showwarning("Missing Input", "Please enter text to find.")
            return False
        return True

    def _on_search_settings_changed(self, *_args):
        """
        Discard on-screen results when the search settings change.

        The results describe offsets found with one folder, search term and case
        setting. Applying them under different settings would validate old
        offsets against new rules, so the preview is invalidated instead.
        """
        if self.is_processing or not self.all_matches:
            return
        self._clear_results()
        self.progress_var.set("Search settings changed - run Preview Changes again")

    def _clear_results(self):
        """Drop the current results and the tree showing them."""
        self.tree.delete(*self.tree.get_children())
        self.match_tree_map.clear()
        self.file_tree_map.clear()
        self.all_matches.clear()
        self.scan_results = []
        self.apply_btn.configure(state=tk.DISABLED)
        self.match_count_label.configure(text="")

    # ------------------------------------------------------------------
    # Preview
    # ------------------------------------------------------------------

    def _start_preview(self):
        """Start the preview scan in a background thread."""
        if not self._validate_inputs():
            return

        self.processing_kind = 'scan'
        self.scan_cancel = threading.Event()

        # Record what this scan is being run with; apply uses these, not the
        # live widgets.
        self.scan_find_text = self.find_var.get()
        self.scan_case_sensitive = self.case_var.get()

        self._clear_results()
        self.progress_var.set("Scanning documents...")
        self.preview_btn.configure(text="Cancel Scan", command=self._cancel_scan)
        self._set_controls_enabled(False)

        thread = threading.Thread(target=self._run_preview, daemon=True)
        thread.start()

    def _cancel_scan(self):
        """Ask the running scan to stop after the file it is on."""
        if self.scan_cancel is not None:
            self.scan_cancel.set()
            self.progress_var.set("Cancelling scan...")
            self.preview_btn.configure(state=tk.DISABLED)

    def _run_preview(self):
        """Background thread: scan documents."""
        folder = self.folder_var.get().strip()
        find_text = self.scan_find_text
        case_sensitive = self.scan_case_sensitive
        cancel = self.scan_cancel

        def progress(file_name, idx, total):
            self.root.after(0, self.progress_var.set,
                            f"Scanning {file_name} ({idx + 1}/{total})...")

        try:
            results, summary = scan_documents_detailed(
                folder, find_text, case_sensitive,
                progress_callback=progress, cancel_event=cancel,
            )
            self.root.after(0, self._display_results, results, summary)
        except Exception as e:
            self.root.after(0, self._preview_error, str(e))

    def _preview_error(self, error_msg: str):
        """Handle preview errors on main thread."""
        self.progress_var.set("Error during scan")
        messagebox.showerror("Scan Error", f"An error occurred:\n{error_msg}")
        self._finish_processing()

    def _display_results(self, results: List[FileResult], summary: dict):
        """Populate the treeview with scan results (main thread)."""
        self.scan_results = results
        self.all_matches.clear()

        total_matches = 0
        total_files = 0

        for fr in results:
            # File-level node
            if fr.error:
                label = f"{WARNING} {fr.file_name} - ERROR: {fr.error}"
                file_node = self.tree.insert("", tk.END, text=label)
            else:
                label = (f"{CHECKED} {fr.file_name} "
                         f"({fr.match_count} match{'es' if fr.match_count != 1 else ''})")
                file_node = self.tree.insert("", tk.END, text=label)
                self.file_tree_map[file_node] = fr
                total_files += 1

                for match in fr.matches:
                    self.all_matches.append(match)
                    total_matches += 1
                    match_node = self.tree.insert(file_node, tk.END,
                                                  text=self._match_label(match))
                    self.match_tree_map[match_node] = match

            # Expand file nodes
            self.tree.item(file_node, open=True)

        scanned = summary.get('files_scanned', 0)
        cancelled = summary.get('cancelled', False)

        if total_matches > 0:
            self.apply_btn.configure(state=tk.NORMAL)
            status = (f"Found {total_matches} match{'es' if total_matches != 1 else ''} "
                      f"in {total_files} file{'s' if total_files != 1 else ''}")
            if cancelled:
                status += f" - scan cancelled after {scanned} file{'s' if scanned != 1 else ''}"
            self.progress_var.set(status)
        elif cancelled:
            self.progress_var.set(
                f"Scan cancelled after {scanned} file{'s' if scanned != 1 else ''}"
            )
        elif scanned == 0:
            self.progress_var.set("No .docx files found in the selected folder")
        else:
            self.progress_var.set(
                f"No matches found in {scanned} .docx file{'s' if scanned != 1 else ''}"
            )

        # Last: this may destroy the window if a close was requested mid-scan.
        self._finish_processing()

    # ------------------------------------------------------------------
    # Results tree
    # ------------------------------------------------------------------

    def _match_label(self, match: Match) -> str:
        """The tree row for one match: checkbox, location, context."""
        check = CHECKED if match.is_selected else UNCHECKED
        prefix = LOCATION_PREFIXES.get(match.location_type, "")
        return f"{check} {prefix}{match.location_detail}: {match.display_context}"

    def _on_tree_click(self, event):
        """Handle clicks to toggle checkboxes."""
        # The worker thread reads the same Match objects, so the tree is inert
        # while a scan or apply is running.
        if self.is_processing:
            return "break"

        # Clicking the expand/collapse arrow must expand, not toggle every
        # checkbox in the file.
        if self.tree.identify_element(event.x, event.y) == "Treeitem.indicator":
            return

        item = self.tree.identify_row(event.y)
        if not item:
            return

        if item in self.match_tree_map:
            match = self.match_tree_map[item]
            match.is_selected = not match.is_selected
            self._refresh_item_label(item, match)
            parent = self.tree.parent(item)
            if parent in self.file_tree_map:
                self._refresh_file_label(parent)

        elif item in self.file_tree_map:
            fr = self.file_tree_map[item]
            # Toggle: if all selected, deselect all; otherwise select all
            new_state = not all(m.is_selected for m in fr.matches)

            for m in fr.matches:
                m.is_selected = new_state

            for child in self.tree.get_children(item):
                if child in self.match_tree_map:
                    self._refresh_item_label(child, self.match_tree_map[child])

            self._refresh_file_label(item)

        self._update_match_count()

    def _refresh_item_label(self, item: str, match: Match):
        """Update a match item's label to reflect its checkbox state."""
        self.tree.item(item, text=self._match_label(match))

    def _refresh_file_label(self, item: str):
        """Update a file item's label to reflect its children's state."""
        fr = self.file_tree_map[item]
        selected = fr.selected_count
        total = fr.match_count

        if selected == total:
            check = CHECKED
        elif selected == 0:
            check = UNCHECKED
        else:
            check = PARTIAL

        label = (f"{check} {fr.file_name} "
                 f"({total} match{'es' if total != 1 else ''}, {selected} selected)")
        self.tree.item(item, text=label)

    def _update_match_count(self):
        """Update the match count label."""
        total = len(self.all_matches)
        selected = sum(1 for m in self.all_matches if m.is_selected)
        self.match_count_label.configure(text=f"{selected} of {total} selected")

        if self.is_processing:
            return
        self.apply_btn.configure(state=tk.NORMAL if selected > 0 else tk.DISABLED)

    def _select_all(self):
        """Select all matches."""
        if self.is_processing:
            return
        for m in self.all_matches:
            m.is_selected = True
        self._refresh_all_labels()
        self._update_match_count()

    def _deselect_all(self):
        """Deselect all matches."""
        if self.is_processing:
            return
        for m in self.all_matches:
            m.is_selected = False
        self._refresh_all_labels()
        self._update_match_count()

    def _refresh_all_labels(self):
        """Refresh all tree labels."""
        for item, match in self.match_tree_map.items():
            self._refresh_item_label(item, match)
        for item in self.file_tree_map:
            self._refresh_file_label(item)

    # ------------------------------------------------------------------
    # Apply
    # ------------------------------------------------------------------

    def _set_controls_enabled(self, enabled: bool):
        """Enable or disable everything that must not be touched mid-run."""
        state = tk.NORMAL if enabled else tk.DISABLED
        for widget in (self.apply_btn, self.browse_btn,
                       self.select_all_btn, self.deselect_all_btn,
                       self.folder_entry, self.find_entry, self.replace_entry):
            widget.configure(state=state)

    def _finish_processing(self):
        """Return the UI to its idle state after a scan or apply."""
        self.processing_kind = None
        self.scan_cancel = None
        self.preview_btn.configure(text="Preview Changes",
                                   command=self._start_preview, state=tk.NORMAL)
        self._set_controls_enabled(True)
        self._update_match_count()

        # The user asked to quit while a scan was running; the scan has now
        # stopped, so honour it.
        if self.close_when_idle:
            self.root.destroy()

    def _start_apply(self):
        """Start applying changes in a background thread."""
        replace_text = self.replace_var.get()
        selected = [m for m in self.all_matches if m.is_selected]
        file_count = len(set(m.file_path for m in selected))

        if not selected:
            messagebox.showinfo("Nothing to Apply", "No changes are selected.")
            return

        confirm = messagebox.askyesno(
            "Confirm Changes",
            f"Apply {len(selected)} replacement{'s' if len(selected) != 1 else ''} "
            f"across {file_count} file{'s' if file_count != 1 else ''}?\n\n"
            # The term actually scanned, which is what these offsets belong to.
            f"Find: \"{self.scan_find_text}\"\n"
            f"Replace: \"{replace_text}\"\n"
            f"Case sensitive: {'yes' if self.scan_case_sensitive else 'no'}\n"
            f"{'Backups will be created.' if self.backup_var.get() else 'NO backups will be created!'}"
        )

        if not confirm:
            return

        self.processing_kind = 'apply'
        self.preview_btn.configure(state=tk.DISABLED)
        self._set_controls_enabled(False)

        thread = threading.Thread(target=self._run_apply, daemon=True)
        thread.start()

    def _run_apply(self):
        """Background thread: apply changes."""
        replace_text = self.replace_var.get()
        create_backups = self.backup_var.get()
        # The setting the offsets were found under, not whatever the checkbox
        # says now.
        case_sensitive = self.scan_case_sensitive

        def progress(file_name, idx, total):
            self.root.after(0, self.progress_var.set,
                           f"Processing {file_name} ({idx + 1}/{total})...")

        try:
            result = apply_changes(
                self.all_matches, replace_text,
                create_backups=create_backups,
                case_sensitive=case_sensitive,
                progress_callback=progress
            )
            self.root.after(0, self._apply_complete, result)
        except Exception as e:
            self.root.after(0, self._apply_error, str(e))

    def _apply_complete(self, result: dict):
        """Handle apply completion on main thread."""
        self.processing_kind = None
        skipped = result.get('total_skipped', 0)

        msg = (f"Completed!\n\n"
               f"Replacements made: {result['total_replaced']}\n"
               f"Files modified: {result['files_modified']}")

        if result.get('backups'):
            msg += f"\nBackups created: {len(result['backups'])}"

        if skipped:
            # A selected match that could not be applied must never disappear
            # quietly - the preview promised it would be changed.
            msg += (f"\n\n{WARNING} {skipped} selected "
                    f"{'change was' if skipped == 1 else 'changes were'} NOT applied.\n"
                    f"The document may have been edited since the preview. "
                    f"Re-run Preview Changes to see the current state.")

        if result['errors']:
            msg += f"\n\nErrors ({len(result['errors'])}):\n"
            for err in result['errors']:
                msg += f"  - {err}\n"

        status = (f"Done: {result['total_replaced']} replacements "
                  f"in {result['files_modified']} files")
        if skipped:
            status += f" ({skipped} skipped)"

        # Results describe the pre-edit documents, so they are stale now.
        self._clear_results()
        self.progress_var.set(status)
        self._finish_processing()

        messagebox.showinfo("Changes Applied", msg)

    def _apply_error(self, error_msg: str):
        """Handle apply errors on main thread."""
        # Keep the results on screen so the run can be retried without
        # re-scanning.
        self.progress_var.set("Error during apply")
        self._finish_processing()
        messagebox.showerror("Apply Error", f"An error occurred:\n{error_msg}")

    # ------------------------------------------------------------------
    # Shutdown
    # ------------------------------------------------------------------

    def _on_close(self):
        """Refuse to close while a worker thread is mid-run."""
        if self.processing_kind == 'apply':
            messagebox.showwarning(
                "Apply in Progress",
                "Documents are being written right now.\n\n"
                "Closing would interrupt the run and leave part of the batch "
                "unmodified. Please wait for it to finish."
            )
            return

        if self.processing_kind == 'scan':
            if not messagebox.askyesno("Scan in Progress",
                                       "A scan is still running. Cancel it and quit?"):
                return
            self.close_when_idle = True
            self._cancel_scan()
            return

        self.root.destroy()


def main():
    # DPI awareness has to be set before the first Tk window exists, otherwise
    # Tk has already sampled the old DPI and the UI stays blurry on high-DPI
    # displays.
    scaling = None
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
        scaling = windll.user32.GetDpiForSystem() / 72.0
    except Exception:
        pass

    root = tk.Tk()

    if scaling:
        try:
            root.tk.call("tk", "scaling", scaling)
        except Exception:
            pass

    FindReplaceApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
