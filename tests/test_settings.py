"""
Rule-set CSV files and remembered application settings.
"""

import json
import os

import pytest

import settings
from document_processor import Rule


def _write(path, text, encoding="utf-8"):
    with open(path, "w", encoding=encoding, newline="") as fh:
        fh.write(text)
    return str(path)


# ---------------------------------------------------------------- rules csv

def test_rules_round_trip_preserves_awkward_text(tmp_path):
    rules = [
        Rule("2022 CBC", "2025 CBC"),
        Rule('say "hi", ok', "multi\nline"),
        Rule(" leading and trailing ", " spaces "),
        Rule("ünïcode – dash", "“quotes”"),
        Rule("NFPA 13, 2022 edition", "NFPA 13, 2025 edition"),
    ]
    path = str(tmp_path / "set.csv")
    assert settings.save_rules_csv(path, rules) == path
    assert settings.load_rules_csv(path) == rules
    with open(path, "rb") as fh:
        assert fh.read().startswith(b"\xef\xbb\xbffind,replace")


def test_header_less_and_bom_less_files_load(tmp_path):
    path = _write(tmp_path / "plain.csv", "2022 CBC,2025 CBC\r\n2022 CFC,2025 CFC\r\n")
    assert settings.load_rules_csv(path) == [Rule("2022 CBC", "2025 CBC"),
                                             Rule("2022 CFC", "2025 CFC")]


def test_header_is_recognised_case_insensitively(tmp_path):
    path = _write(tmp_path / "h.csv", "Find, Replace\n2022,2025\n")
    assert settings.load_rules_csv(path) == [Rule("2022", "2025")]


def test_blank_rows_and_missing_replace_column(tmp_path):
    path = _write(tmp_path / "b.csv", "find,replace\n\n,\n2022\n\n")
    assert settings.load_rules_csv(path) == [Rule("2022", "")]


def test_empty_find_with_a_replacement_is_an_error_naming_the_row(tmp_path):
    path = _write(tmp_path / "e.csv", "find,replace\n2022,2025\n,oops\n")
    with pytest.raises(ValueError) as info:
        settings.load_rules_csv(path)
    assert "Row 3" in str(info.value)


def test_duplicate_find_is_an_error_naming_both_rows(tmp_path):
    path = _write(tmp_path / "d.csv", "find,replace\n2022,2025\nx,y\n2022,2026\n")
    with pytest.raises(ValueError) as info:
        settings.load_rules_csv(path)
    assert "Row 4" in str(info.value) and "row 2" in str(info.value)


def test_extra_columns_are_ignored(tmp_path):
    path = _write(tmp_path / "x.csv", "find,replace,note\n2022,2025,code cycle\n")
    assert settings.load_rules_csv(path) == [Rule("2022", "2025")]


# ---------------------------------------------------------------- settings path

def test_default_path_uses_appdata(monkeypatch, tmp_path):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    assert settings.default_settings_path() == os.path.join(
        str(tmp_path), "DocxFindReplace", "settings.json")


def test_default_path_falls_back_to_home(monkeypatch, tmp_path):
    monkeypatch.delenv("APPDATA", raising=False)
    monkeypatch.setenv("HOME", str(tmp_path))
    monkeypatch.setenv("USERPROFILE", str(tmp_path))
    assert settings.default_settings_path() == os.path.join(
        str(tmp_path), ".DocxFindReplace", "settings.json")


# ---------------------------------------------------------------- load / save

def test_missing_file_gives_defaults(tmp_path):
    loaded = settings.load_settings(str(tmp_path / "none.json"))
    assert loaded == {**settings.DEFAULTS, "rules": []}
    assert loaded["rules"] is not settings.DEFAULTS["rules"]


def test_corrupt_file_gives_defaults(tmp_path):
    path = _write(tmp_path / "bad.json", "{not json")
    assert settings.load_settings(path) == {**settings.DEFAULTS, "rules": []}
    path = _write(tmp_path / "list.json", "[1, 2]")
    assert settings.load_settings(path) == {**settings.DEFAULTS, "rules": []}


def test_wrong_types_and_unknown_keys_are_ignored(tmp_path):
    data = {"folder": 12, "case_sensitive": "yes", "whole_word": 1, "recursive": True,
            "author": "Abe", "mystery": True,
            "rules": [{"find": "a", "replace": "b"}, {"find": "", "replace": "x"},
                      "junk", {"replace": "only"}, {"find": "c"}]}
    path = _write(tmp_path / "s.json", json.dumps(data))
    loaded = settings.load_settings(path)
    assert loaded["folder"] == ""
    assert loaded["case_sensitive"] is True
    assert loaded["whole_word"] is False
    assert loaded["recursive"] is True
    assert loaded["author"] == "Abe"
    assert "mystery" not in loaded
    assert loaded["rules"] == [Rule("a", "b"), Rule("c", "")]


def test_save_creates_the_folder_and_round_trips_rules(tmp_path):
    path = str(tmp_path / "deep" / "er" / "settings.json")
    data = {"folder": "C:\\specs", "rules": [Rule("2022 CBC", "2025 CBC")],
            "case_sensitive": False, "whole_word": True, "recursive": True,
            "create_backups": False, "write_change_log": False,
            "tracked_changes": True, "author": "Abe", "geometry": "900x700+10+10",
            "ignored": "x"}
    assert settings.save_settings(data, path) == path
    assert not [f for f in os.listdir(os.path.dirname(path)) if f.endswith(".tmp")]
    loaded = settings.load_settings(path)
    expected = {k: v for k, v in data.items() if k != "ignored"}
    assert loaded == expected
    with open(path, encoding="utf-8") as fh:
        raw = json.load(fh)
    assert raw["rules"] == [{"find": "2022 CBC", "replace": "2025 CBC"}]
    assert "ignored" not in raw


def test_save_accepts_rule_dicts(tmp_path):
    path = str(tmp_path / "settings.json")
    settings.save_settings({"rules": [{"find": "a", "replace": "b"}]}, path)
    assert settings.load_settings(path)["rules"] == [Rule("a", "b")]


def test_save_overwrites_atomically(tmp_path):
    path = str(tmp_path / "settings.json")
    settings.save_settings({"author": "first"}, path)
    settings.save_settings({"author": "second"}, path)
    assert settings.load_settings(path)["author"] == "second"
    assert sorted(os.listdir(str(tmp_path))) == ["settings.json"]
