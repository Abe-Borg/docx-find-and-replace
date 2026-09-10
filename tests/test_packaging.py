"""
Packaging guards.

The build config itself cannot run here - PyInstaller and Inno Setup are
Windows-only and CI is the real check - but the release-tag guard is ordinary
Python, so it is tested like anything else rather than trusted because it looks
right in a YAML file.
"""

import os
import subprocess
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "packaging"))

import check_tag  # noqa: E402
import version  # noqa: E402


# ----------------------------------------------------- tag normalisation

@pytest.mark.parametrize("tag,expected", [
    ("v1.0.0", "1.0.0"),
    ("V1.0.0", "1.0.0"),
    ("1.0.0", "1.0.0"),
    ("  v1.0.0  ", "1.0.0"),
    ("v1.0.0-rc1", "1.0.0-rc1"),
])
def test_normalize_strips_one_leading_v(tag, expected):
    assert check_tag.normalize(tag) == expected


def test_normalize_leaves_an_inner_v_alone():
    assert check_tag.normalize("v1.0.0v") == "1.0.0v"


# ----------------------------------------------------- the guard itself

def test_matching_tag_passes():
    assert check_tag.check("v1.0.0", "1.0.0") == ""


def test_tag_ahead_of_version_py_is_rejected():
    """
    The case that would have shipped 1.0.0 binaries inside a v1.1.0 release.
    """
    problem = check_tag.check("v1.1.0", "1.0.0")
    assert problem
    assert "does not match version.py" in problem
    assert "1.1.0" in problem and "1.0.0" in problem


def test_tag_behind_version_py_is_rejected():
    assert check_tag.check("v0.9.0", "1.0.0")


def test_prerelease_tag_must_match_exactly():
    assert check_tag.check("v1.0.0-rc1", "1.0.0")
    assert check_tag.check("v1.0.0-rc1", "1.0.0-rc1") == ""


def test_empty_tag_is_rejected():
    assert check_tag.check("   ", "1.0.0")


def test_check_defaults_to_the_real_version():
    assert check_tag.check(f"v{version.__version__}") == ""
    assert check_tag.check("v999.0.0")


# ----------------------------------------------------- as the workflow runs it

def _run(*args):
    return subprocess.run(
        [sys.executable, os.path.join(ROOT, "packaging", "check_tag.py"), *args],
        capture_output=True, text=True,
    )


def test_cli_exits_zero_on_a_matching_tag():
    result = _run(f"v{version.__version__}")
    assert result.returncode == 0, result.stderr
    assert "matches version.py" in result.stdout


def test_cli_exits_nonzero_on_a_mismatched_tag():
    result = _run("v99.9.9")
    assert result.returncode == 1
    assert "does not match version.py" in result.stderr


def test_cli_without_a_tag_is_a_usage_error():
    assert _run().returncode == 2
