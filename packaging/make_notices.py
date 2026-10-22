"""
Write THIRD-PARTY-NOTICES.txt: the license of everything the packaged build carries.

The executables bundle the Python runtime (with the libraries the Windows
build of Python links, Tcl/Tk among them), every package requirements.txt
pulls in, and PyInstaller's bootloader. Their licenses ask for the copyright
and license notice to travel with binary copies, so the build bundles this
file, the installer installs it, and Help > Third-Party Notices shows it.

Generated from the installed packages' own metadata at build time rather than
written by hand, so a dependency added to requirements.txt is picked up and a
version bump carries its own license text. lxml's license file covers the C
libraries its wheels link statically (libxml2, libxslt, zlib, iconv). A
component whose license text cannot be found fails the build rather than being
left out of the file.

Usage:  python packaging/make_notices.py [output]
Writes: packaging/THIRD-PARTY-NOTICES.txt
"""

import glob
import importlib.metadata as metadata
import os
import re
import sys
import sysconfig
from typing import Dict, Iterable, List, NamedTuple, Optional, Sequence, Tuple

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

from version import APP_NAME, __version__  # noqa: E402

DEFAULT_OUTPUT = os.path.join(HERE, "THIRD-PARTY-NOTICES.txt")

# Build tools whose code ends up in the exe, by distribution name, with the name
# and license the notice gives them. Only PyInstaller's bootloader is bundled,
# not the tool or its dependencies, and the bootloader is under the GPL's
# bootloader exception, which places no condition on the program it starts.
BUILD_TOOLS = {
    "pyinstaller": ("PyInstaller bootloader", "GPL-2.0-or-later WITH Bootloader-exception"),
}

_LICENSE_NAME = re.compile(r"(LICEN[CS]E|COPYING|NOTICE)", re.IGNORECASE)
_REQUIREMENT_NAME = re.compile(r"^\s*([A-Za-z0-9][A-Za-z0-9._-]*)")
_EXTRA_MARKER = re.compile(r";.*\bextra\s*==")


class Component(NamedTuple):
    name: str
    version: str
    license_name: str
    url: str
    texts: List[Tuple[str, str]]      # (label, license text)


class NoticeError(Exception):
    pass


def _canonical(name: str) -> str:
    return re.sub(r"[-_.]+", "-", name).lower()


def requirement_names(lines: Iterable[str]) -> List[str]:
    """Package names in requirements.txt lines, skipping comments and options."""
    names = []
    for line in lines:
        line = line.split("#", 1)[0].strip()
        if not line or line.startswith("-"):
            continue
        m = _REQUIREMENT_NAME.match(line)
        if m:
            names.append(m.group(1))
    return names


def runtime_closure(roots: Sequence[str]) -> List[str]:
    """
    The installed distributions `roots` need at run time, roots included.

    Requirements behind an extra are not followed: nothing installs the extras.
    Raises NoticeError for a requirement that is not installed, because a build
    cannot bundle what is not there and the notices must not guess.
    """
    seen, order, queue = set(), [], list(roots)
    while queue:
        name = queue.pop(0)
        key = _canonical(name)
        if key in seen:
            continue
        seen.add(key)
        try:
            dist = metadata.distribution(name)
        except metadata.PackageNotFoundError:
            raise NoticeError(f"{name} is required but not installed")
        order.append(dist.metadata["Name"])
        for req in dist.requires or []:
            if _EXTRA_MARKER.search(req):
                continue
            m = _REQUIREMENT_NAME.match(req)
            if m:
                queue.append(m.group(1))
    return order


def _license_name(meta) -> str:
    expression = meta.get("License-Expression")
    if expression:
        return expression
    declared = (meta.get("License") or "").strip()
    if declared and "\n" not in declared and len(declared) <= 60:
        return declared
    for classifier in meta.get_all("Classifier") or []:
        if classifier.startswith("License :: "):
            return classifier.rsplit(" :: ", 1)[-1]
    return "see license text below"


def _url(meta) -> str:
    for entry in meta.get_all("Project-URL") or []:
        label, _, url = entry.partition(",")
        if label.strip().lower() in ("homepage", "home", "source", "source code",
                                     "repository"):
            return url.strip()
    return (meta.get("Home-page") or "").strip() or \
        f"https://pypi.org/project/{meta['Name']}/"


def distribution_component(name: str) -> Component:
    """A package's notice, from the license files in its installed metadata."""
    try:
        dist = metadata.distribution(name)
    except metadata.PackageNotFoundError:
        raise NoticeError(f"{name} is not installed")
    texts = []
    for f in sorted(dist.files or [], key=str):
        base = os.path.basename(str(f))
        if not _LICENSE_NAME.search(base) or base.endswith((".py", ".pyc")):
            continue
        if ".dist-info" not in str(f) and ".egg-info" not in str(f):
            continue
        path = dist.locate_file(f)
        try:
            with open(path, encoding="utf-8", errors="replace") as fh:
                texts.append((base, fh.read().strip()))
        except OSError:
            continue
    meta = dist.metadata
    if not texts:
        raise NoticeError(f"no license file found in the metadata of {meta['Name']}")
    return Component(meta["Name"], dist.version, _license_name(meta), _url(meta), texts)


def python_component(license_paths: Optional[Sequence[str]] = None,
                     tcl_dir: Optional[str] = None) -> Component:
    """
    The Python runtime. Its Windows LICENSE.txt also carries the licenses of
    the libraries that build links (OpenSSL, libffi, bzip2, xz, zlib, ...);
    Tcl/Tk's own license.terms are added when present.
    """
    if license_paths is None:
        license_paths = [os.path.join(sys.base_prefix, "LICENSE.txt"),
                         os.path.join(sys.base_prefix, "LICENSE"),
                         os.path.join(sysconfig.get_path("stdlib"), "LICENSE.txt")]
    if tcl_dir is None:
        tcl_dir = os.path.join(sys.base_prefix, "tcl")
    texts = []
    for path in license_paths:
        if os.path.isfile(path):
            with open(path, encoding="utf-8", errors="replace") as fh:
                texts.append(("Python license", fh.read().strip()))
            break
    if not texts:
        raise NoticeError("the Python license was not found (looked in "
                          + ", ".join(license_paths) + ")")
    for terms in sorted(glob.glob(os.path.join(tcl_dir, "*", "license.terms"))):
        with open(terms, encoding="utf-8", errors="replace") as fh:
            label = os.path.basename(os.path.dirname(terms))
            texts.append((f"{label} license.terms", fh.read().strip()))
    return Component("Python", sys.version.split()[0], "PSF-2.0",
                     "https://www.python.org/", texts)


def collect(roots: Sequence[str], build_tools: Optional[Dict[str, Tuple[str, str]]] = None,
            python: Optional[Component] = None) -> List[Component]:
    build_tools = BUILD_TOOLS if build_tools is None else build_tools
    components = [python or python_component()]
    names = runtime_closure(roots)
    components += [distribution_component(n) for n in sorted(names, key=str.lower)]
    for dist_name, (shown_name, license_name) in build_tools.items():
        components.append(distribution_component(dist_name)._replace(
            name=shown_name, license_name=license_name))
    return components


def render(components: Sequence[Component]) -> str:
    rule = "=" * 78
    lines = [
        "THIRD-PARTY SOFTWARE NOTICES",
        f"{APP_NAME} {__version__}",
        "",
        f"{APP_NAME} is licensed under its own End User License Agreement",
        "(LICENSE.txt). It includes the third-party components below, each licensed",
        "under its own terms, reproduced in full. Nothing in the End User License",
        "Agreement restricts the rights these licenses grant in the components",
        "they cover. Source code for each component is available from the address",
        "listed for it.",
        "",
        "Components:",
    ]
    for c in components:
        lines.append(f"  - {c.name} {c.version} ({c.license_name})  {c.url}")
    for c in components:
        lines += ["", rule, f"{c.name} {c.version}", c.url, rule]
        for label, text in c.texts:
            lines += ["", f"--- {label} ---", "", text]
    return "\n".join(lines) + "\n"


def main(argv: Sequence[str]) -> int:
    out = argv[1] if len(argv) > 1 else DEFAULT_OUTPUT
    with open(os.path.join(ROOT, "requirements.txt"), encoding="utf-8") as fh:
        roots = requirement_names(fh)
    try:
        components = collect(roots)
    except NoticeError as e:
        print(f"make_notices: {e}", file=sys.stderr)
        return 1
    with open(out, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(render(components))
    print(f"wrote {out}: " + ", ".join(f"{c.name} {c.version}" for c in components))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
