"""
main.py
Word Document Batch Find & Replace - GUI Application

A tkinter-based GUI for performing batch find-and-replace operations
across multiple Word documents (.docx) with a preview-before-commit workflow.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from document_processor import scan_documents, apply_changes, FileResult, Match
from typing import List


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
        self.match_tree_map = {}  # tree item id -> Match object
        self.file_tree_map = {}   # tree item id -> FileResult object
        self.has_previewed = False
        self.is_processing = False

        # Configure styles
        style = ttk.Style()
        style.configure("TButton", padding=4)
        style.configure("Header.TLabel", font=("Segoe UI", 10, "bold"))
        style.configure("Status.TLabel", font=("Segoe UI", 9))

        self._build_ui()

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
        ttk.Button(folder_frame, text="Browse...", command=self._browse_folder).pack(side=tk.LEFT)

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
        ttk.Checkbutton(options_frame, text="Create backups (.docx.bak)",
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
        ttk.Button(toolbar, text="Select All", command=self._select_all).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(toolbar, text="Deselect All", command=self._deselect_all).pack(side=tk.LEFT)
        self.match_count_label = ttk.Label(toolbar, text="", style="Status.TLabel")
        self.match_count_label.pack(side=tk.RIGHT)

        # Treeview with checkboxes
        tree_frame = ttk.Frame(results_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(tree_frame, columns=("status", "context"), show="tree",
                                  selectmode="browse")
        self.tree.column("#0", width=350, minwidth=200)
        self.tree.column("status", width=0, stretch=False)
        self.tree.column("context", width=0, stretch=False)

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

    def _start_preview(self):
        """Start the preview scan in a background thread."""
        if not self._validate_inputs():
            return

        self.is_processing = True
        self.preview_btn.configure(state=tk.DISABLED)
        self.apply_btn.configure(state=tk.DISABLED)
        self.progress_var.set("Scanning documents...")
        self.tree.delete(*self.tree.get_children())
        self.match_tree_map.clear()
        self.file_tree_map.clear()
        self.all_matches.clear()

        thread = threading.Thread(target=self._run_preview, daemon=True)
        thread.start()

    def _run_preview(self):
        """Background thread: scan documents."""
        folder = self.folder_var.get().strip()
        find_text = self.find_var.get()
        case_sensitive = self.case_var.get()

        try:
            results = scan_documents(folder, find_text, case_sensitive)
            self.root.after(0, self._display_results, results)
        except Exception as e:
            self.root.after(0, self._preview_error, str(e))

    def _preview_error(self, error_msg: str):
        """Handle preview errors on main thread."""
        self.is_processing = False
        self.preview_btn.configure(state=tk.NORMAL)
        self.progress_var.set("Error during scan")
        messagebox.showerror("Scan Error", f"An error occurred:\n{error_msg}")

    def _display_results(self, results: List[FileResult]):
        """Populate the treeview with scan results (main thread)."""
        self.scan_results = results
        self.all_matches.clear()
        self.has_previewed = True
        self.is_processing = False
        self.preview_btn.configure(state=tk.NORMAL)

        total_matches = 0
        total_files = 0

        for fr in results:
            # File-level node
            if fr.error:
                label = f"\u26A0 {fr.file_name} - ERROR: {fr.error}"
                file_node = self.tree.insert("", tk.END, text=label)
            else:
                check = "\u2611"  # ☑
                label = f"{check} {fr.file_name} ({fr.match_count} match{'es' if fr.match_count != 1 else ''})"
                file_node = self.tree.insert("", tk.END, text=label)
                self.file_tree_map[file_node] = fr
                total_files += 1

                for match in fr.matches:
                    self.all_matches.append(match)
                    total_matches += 1

                    loc_prefix = ""
                    if match.location_type == 'table':
                        loc_prefix = f"[Table] "
                    elif match.location_type == 'header':
                        loc_prefix = f"[Header] "
                    elif match.location_type == 'footer':
                        loc_prefix = f"[Footer] "

                    m_check = "\u2611" if match.is_selected else "\u2610"
                    m_label = f"{m_check} {loc_prefix}{match.location_detail}: {match.display_context}"
                    match_node = self.tree.insert(file_node, tk.END, text=m_label)
                    self.match_tree_map[match_node] = match

            # Expand file nodes
            self.tree.item(file_node, open=True)

        if total_matches > 0:
            self.apply_btn.configure(state=tk.NORMAL)
            self.progress_var.set(f"Found {total_matches} match{'es' if total_matches != 1 else ''} "
                                   f"in {total_files} file{'s' if total_files != 1 else ''}")
        elif not results:
            self.progress_var.set("No .docx files found in the selected folder")
        else:
            self.progress_var.set("No matches found")

        self._update_match_count()

    def _on_tree_click(self, event):
        """Handle clicks to toggle checkboxes."""
        item = self.tree.identify_row(event.y)
        if not item:
            return

        # Check if it's a match-level item
        if item in self.match_tree_map:
            match = self.match_tree_map[item]
            match.is_selected = not match.is_selected
            self._refresh_item_label(item, match)
            # Update parent file node
            parent = self.tree.parent(item)
            if parent in self.file_tree_map:
                self._refresh_file_label(parent)

        # Check if it's a file-level item
        elif item in self.file_tree_map:
            fr = self.file_tree_map[item]
            # Toggle: if all selected, deselect all; otherwise select all
            all_selected = all(m.is_selected for m in fr.matches)
            new_state = not all_selected

            for m in fr.matches:
                m.is_selected = new_state

            # Refresh all child items
            for child in self.tree.get_children(item):
                if child in self.match_tree_map:
                    self._refresh_item_label(child, self.match_tree_map[child])

            self._refresh_file_label(item)

        self._update_match_count()

    def _refresh_item_label(self, item: str, match: Match):
        """Update a match item's label to reflect its checkbox state."""
        check = "\u2611" if match.is_selected else "\u2610"

        loc_prefix = ""
        if match.location_type == 'table':
            loc_prefix = "[Table] "
        elif match.location_type == 'header':
            loc_prefix = "[Header] "
        elif match.location_type == 'footer':
            loc_prefix = "[Footer] "

        label = f"{check} {loc_prefix}{match.location_detail}: {match.display_context}"
        self.tree.item(item, text=label)

    def _refresh_file_label(self, item: str):
        """Update a file item's label to reflect its children's state."""
        fr = self.file_tree_map[item]
        selected = fr.selected_count
        total = fr.match_count

        if selected == total:
            check = "\u2611"
        elif selected == 0:
            check = "\u2610"
        else:
            check = "\u2612"  # ☒ (partial)

        label = f"{check} {fr.file_name} ({total} match{'es' if total != 1 else ''}, {selected} selected)"
        self.tree.item(item, text=label)

    def _update_match_count(self):
        """Update the match count label."""
        total = len(self.all_matches)
        selected = sum(1 for m in self.all_matches if m.is_selected)
        self.match_count_label.configure(text=f"{selected} of {total} selected")

        if selected > 0:
            self.apply_btn.configure(state=tk.NORMAL)
        else:
            self.apply_btn.configure(state=tk.DISABLED)

    def _select_all(self):
        """Select all matches."""
        for m in self.all_matches:
            m.is_selected = True
        self._refresh_all_labels()
        self._update_match_count()

    def _deselect_all(self):
        """Deselect all matches."""
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

    def _start_apply(self):
        """Start applying changes in a background thread."""
        replace_text = self.replace_var.get()
        selected_count = sum(1 for m in self.all_matches if m.is_selected)
        file_count = len(set(m.file_path for m in self.all_matches if m.is_selected))

        if selected_count == 0:
            messagebox.showinfo("Nothing to Apply", "No changes are selected.")
            return

        confirm = messagebox.askyesno(
            "Confirm Changes",
            f"Apply {selected_count} replacement{'s' if selected_count != 1 else ''} "
            f"across {file_count} file{'s' if file_count != 1 else ''}?\n\n"
            f"Find: \"{self.find_var.get()}\"\n"
            f"Replace: \"{replace_text}\"\n"
            f"{'Backups will be created.' if self.backup_var.get() else 'NO backups will be created!'}"
        )

        if not confirm:
            return

        self.is_processing = True
        self.apply_btn.configure(state=tk.DISABLED)
        self.preview_btn.configure(state=tk.DISABLED)

        thread = threading.Thread(target=self._run_apply, daemon=True)
        thread.start()

    def _run_apply(self):
        """Background thread: apply changes."""
        replace_text = self.replace_var.get()
        create_backups = self.backup_var.get()
        case_sensitive = self.case_var.get()

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
        self.is_processing = False
        self.preview_btn.configure(state=tk.NORMAL)
        self.has_previewed = False

        msg = (f"Completed!\n\n"
               f"Replacements made: {result['total_replaced']}\n"
               f"Files modified: {result['files_modified']}")

        if result['errors']:
            msg += f"\n\nErrors ({len(result['errors'])}):\n"
            for err in result['errors']:
                msg += f"  - {err}\n"

        self.progress_var.set(
            f"Done: {result['total_replaced']} replacements in {result['files_modified']} files"
        )

        messagebox.showinfo("Changes Applied", msg)

        # Clear the tree since results are now stale
        self.tree.delete(*self.tree.get_children())
        self.match_tree_map.clear()
        self.file_tree_map.clear()
        self.all_matches.clear()
        self.apply_btn.configure(state=tk.DISABLED)

    def _apply_error(self, error_msg: str):
        """Handle apply errors on main thread."""
        self.is_processing = False
        self.preview_btn.configure(state=tk.NORMAL)
        self.progress_var.set("Error during apply")
        messagebox.showerror("Apply Error", f"An error occurred:\n{error_msg}")


def main():
    root = tk.Tk()

    # Set DPI awareness for crisp text on Windows
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = FindReplaceApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
