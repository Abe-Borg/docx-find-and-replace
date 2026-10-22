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


# ----------------------------------------------------- third-party notices

import make_notices  # noqa: E402


def _read(*parts):
    with open(os.path.join(ROOT, *parts), encoding="utf-8") as fh:
        return fh.read()


def _python(tmp_path):
    lic = tmp_path / "LICENSE.txt"
    lic.write_text("PYTHON SOFTWARE FOUNDATION LICENSE VERSION 2", encoding="utf-8")
    return make_notices.python_component([str(lic)], tcl_dir=str(tmp_path / "no-tcl"))


def test_requirement_names_skip_comments_options_and_versions():
    lines = ["# runtime", "-r other.txt", "python-docx>=1.1.0,<2.0  # pinned",
             "", "lxml ; sys_platform == 'win32'"]
    assert make_notices.requirement_names(lines) == ["python-docx", "lxml"]


def test_the_runtime_closure_follows_python_docx_to_what_it_needs():
    names = {make_notices._canonical(n) for n in make_notices.runtime_closure(["python-docx"])}
    assert {"python-docx", "lxml", "typing-extensions"} <= names


def test_a_missing_requirement_fails_rather_than_being_left_out():
    with pytest.raises(make_notices.NoticeError):
        make_notices.runtime_closure(["no-such-package-xyz"])


def test_notices_carry_each_bundled_license_in_full(tmp_path):
    components = make_notices.collect(["python-docx"], build_tools={},
                                      python=_python(tmp_path))
    text = make_notices.render(components)
    assert "PYTHON SOFTWARE FOUNDATION LICENSE VERSION 2" in text
    for name in ("python-docx", "lxml"):
        component = next(c for c in components if c.name == name)
        assert f"{component.name} {component.version}" in text
        for _label, license_text in component.texts:
            assert license_text in text
    # lxml's license file is what covers the C libraries its wheels link.
    assert "libxml2" in text
    assert version.APP_NAME in text and version.__version__ in text


def test_a_build_tool_is_named_as_the_part_that_is_bundled(tmp_path):
    components = make_notices.collect(
        [], build_tools={"python-docx": ("python-docx core", "MIT")}, python=_python(tmp_path))
    assert components[-1].name == "python-docx core"
    assert components[-1].texts


def test_the_python_license_is_required(tmp_path):
    with pytest.raises(make_notices.NoticeError):
        make_notices.python_component([str(tmp_path / "absent.txt")])


def test_tcl_and_tk_license_terms_are_added_when_present(tmp_path):
    for lib in ("tcl8.6", "tk8.6"):
        (tmp_path / "tcl" / lib).mkdir(parents=True)
        (tmp_path / "tcl" / lib / "license.terms").write_text(f"{lib} terms", encoding="utf-8")
    lic = tmp_path / "LICENSE.txt"
    lic.write_text("PSF", encoding="utf-8")
    component = make_notices.python_component([str(lic)], tcl_dir=str(tmp_path / "tcl"))
    assert [label for label, _ in component.texts] == [
        "Python license", "tcl8.6 license.terms", "tk8.6 license.terms"]


# ----------------------------------------------------- attribution stays in step

def test_license_carries_the_copyright_and_contact_from_version_py():
    text = _read("LICENSE.txt")
    assert version.COPYRIGHT in text
    assert version.SUPPORT_URL in text
    assert version.RELEASES_URL in text


def _iss_defines():
    defines = {}
    for line in _read("packaging", "installer.iss").splitlines():
        parts = line.split(None, 2)
        if len(parts) == 3 and parts[0] == "#define" and parts[2].startswith('"'):
            defines[parts[1]] = parts[2].strip().strip('"')
    return defines


def test_installer_attribution_matches_version_py():
    defines = _iss_defines()
    assert defines["AppPublisher"] == version.PUBLISHER
    assert defines["AppCopyright"] == version.COPYRIGHT
    assert defines["AppURL"] == version.HOMEPAGE_URL
    assert defines["AppId"] == version.APP_ID


def test_installer_shows_and_installs_the_license_and_notices():
    iss = _read("packaging", "installer.iss")
    assert "LicenseFile=..\\LICENSE.txt" in iss
    assert 'Source: "..\\LICENSE.txt"' in iss
    assert 'Source: "THIRD-PARTY-NOTICES.txt"' in iss
    assert "AppContact=" in iss and "AppSupportURL=" in iss


@pytest.mark.parametrize("spec", ["app.spec", "app-portable.spec"])
def test_both_builds_bundle_the_license_and_notices(spec):
    text = _read("packaging", spec)
    assert '"LICENSE.txt"' in text
    assert '"THIRD-PARTY-NOTICES.txt"' in text
