"""
settings.py
Rule-set files and remembered application settings.

Two plain-file formats, both deliberately simple enough to edit by hand:

* A **rule set** is a CSV with a ``find,replace`` header and one rule per
  row - the twelve substitutions of a code-cycle update, kept from project
  to project. Written as UTF-8 with a byte-order mark so Excel round-trips
  it; read tolerantly (with or without the mark, with or without the header).

* **Settings** are a JSON object under ``%APPDATA%\\DocxFindReplace`` (or
  ``~/.DocxFindReplace`` when APPDATA is unset) holding the last folder,
  rules and options so the tool opens the way it was closed. A missing or
  corrupt file yields the defaults; unknown keys are ignored and a value of
  the wrong type is dropped, so an old or edited file can never crash the
  start-up.

No tkinter here: this module is imported by the GUI and tested headlessly.
"""

import csv
import json
import os
from typing import Any, Dict, List, Optional, Sequence

from document_processor import Rule

APP_DIR_NAME = "DocxFindReplace"
SETTINGS_FILE_NAME = "settings.json"

RULES_HEADER = ("find", "replace")

DEFAULTS: Dict[str, Any] = {
    "folder": "",
    "rules": [],                  # list of {"find": str, "replace": str}
    "case_sensitive": True,
    "whole_word": False,
    "recursive": False,
    "create_backups": True,
    "write_change_log": True,
    "tracked_changes": False,
    "author": "",
    "geometry": "",
}


# --------------------------------------------------------------------------
# Rule sets (CSV)
# --------------------------------------------------------------------------

def load_rules_csv(path: str) -> List[Rule]:
    """
    Read a rule set. Blank rows are skipped; a header row is optional.

    Fields are taken exactly as written - a rule may legitimately start or
    end with a space. Raises ``ValueError`` naming the row for a rule whose
    search text is empty but has a replacement, and for two rules with the
    same search text.
    """
    rules: List[Rule] = []
    seen: Dict[str, int] = {}
    with open(path, encoding="utf-8-sig", newline="") as fh:
        for row_no, row in enumerate(csv.reader(fh), start=1):
            if not row or all(cell == "" for cell in row):
                continue
            find_text = row[0]
            replace_text = row[1] if len(row) > 1 else ""
            if (row_no == 1 and find_text.strip().lower() == RULES_HEADER[0]
                    and replace_text.strip().lower() == RULES_HEADER[1]):
                continue
            if not find_text:
                raise ValueError(f"Row {row_no}: the find text is empty.")
            if find_text in seen:
                raise ValueError(
                    f'Row {row_no}: "{find_text}" is already a rule (row {seen[find_text]}).')
            seen[find_text] = row_no
            rules.append(Rule(find_text, replace_text))
    return rules


def save_rules_csv(path: str, rules: Sequence[Rule]) -> str:
    """Write a rule set with a header row. Returns `path`."""
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(RULES_HEADER)
        for rule in rules:
            writer.writerow([rule.find_text, rule.replace_text])
    return path


# --------------------------------------------------------------------------
# Application settings (JSON)
# --------------------------------------------------------------------------

def default_settings_path() -> str:
    base = os.environ.get("APPDATA")
    if base:
        return os.path.join(base, APP_DIR_NAME, SETTINGS_FILE_NAME)
    return os.path.join(os.path.expanduser("~"), f".{APP_DIR_NAME}", SETTINGS_FILE_NAME)


def _rules_from_json(value: Any) -> List[Rule]:
    rules: List[Rule] = []
    if not isinstance(value, list):
        return rules
    for item in value:
        if not isinstance(item, dict):
            continue
        find_text = item.get("find")
        replace_text = item.get("replace", "")
        if isinstance(find_text, str) and find_text and isinstance(replace_text, str):
            rules.append(Rule(find_text, replace_text))
    return rules


def load_settings(path: Optional[str] = None) -> Dict[str, Any]:
    """
    The remembered settings, with defaults filled in.

    ``rules`` comes back as a list of `Rule`. Anything that cannot be read
    or does not look like the expected type falls back to its default.
    """
    path = path or default_settings_path()
    settings: Dict[str, Any] = dict(DEFAULTS)
    settings["rules"] = []
    try:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)
    except (OSError, ValueError):
        return settings
    if not isinstance(data, dict):
        return settings

    for key, default in DEFAULTS.items():
        if key not in data:
            continue
        value = data[key]
        if key == "rules":
            settings["rules"] = _rules_from_json(value)
        elif isinstance(default, bool):
            if isinstance(value, bool):
                settings[key] = value
        elif isinstance(value, type(default)):
            settings[key] = value
    return settings


def save_settings(settings: Dict[str, Any], path: Optional[str] = None) -> str:
    """
    Write settings, creating the folder if needed. Returns the path written.

    `settings["rules"]` may hold `Rule` objects or ``{"find", "replace"}``
    dicts. Written through a temporary file so a crash mid-write leaves the
    previous file intact.
    """
    path = path or default_settings_path()
    data: Dict[str, Any] = {}
    for key in DEFAULTS:
        if key not in settings:
            continue
        value = settings[key]
        if key == "rules":
            value = [
                {"find": r.find_text, "replace": r.replace_text} if isinstance(r, Rule)
                else {"find": r.get("find", ""), "replace": r.get("replace", "")}
                for r in value
            ]
        data[key] = value

    directory = os.path.dirname(os.path.abspath(path))
    os.makedirs(directory, exist_ok=True)
    tmp_path = os.path.join(directory, f".{os.path.basename(path)}.{os.getpid()}.tmp")
    try:
        with open(tmp_path, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=2, ensure_ascii=False)
        os.replace(tmp_path, path)
    except Exception:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
        raise
    return path
