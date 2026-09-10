"""
Fail a release build when the git tag disagrees with version.py.

Everything that names or stamps an artifact - the installer filename, the
executable's Windows file-version resource, the uploaded artifact - comes from
version.py. The release, however, is created from the pushed tag. Nothing
reconciles the two, so pushing v1.1.0 while version.py still said 1.0.0 would
publish a release labelled 1.1.0 containing binaries stamped 1.0.0, wrong in
both their filenames and their file properties.

A tag is a promise about what is inside the artifacts. This check makes the
build fail before publishing rather than quietly break that promise.

Usage:  python packaging/check_tag.py v1.2.3
Exits:  0 when the tag matches version.py, 1 when it does not.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from version import __version__  # noqa: E402


def normalize(tag: str) -> str:
    """Strip a single leading 'v' and surrounding whitespace from a tag."""
    tag = tag.strip()
    return tag[1:] if tag[:1].lower() == "v" else tag


def check(tag: str, version: str = None) -> str:
    """
    Return an error message when `tag` disagrees with the version, else "".

    Prerelease tags are held to the same rule on purpose: a v1.0.0-rc1 release
    should be built from a version.py that also says 1.0.0-rc1, otherwise the
    binaries inside it are indistinguishable from the final ones.
    """
    version = __version__ if version is None else version
    if not tag.strip():
        return "No tag was supplied to check against version.py."
    normalized = normalize(tag)
    if normalized != version:
        return (
            f"Tag {tag.strip()} does not match version.py ({version}). "
            f"Bump __version__ in version.py to {normalized} and push that "
            f"first, or delete the tag and retag the corrected commit."
        )
    return ""


def main(argv) -> int:
    if len(argv) < 2:
        print("usage: check_tag.py <tag>", file=sys.stderr)
        return 2
    problem = check(argv[1])
    if problem:
        print(problem, file=sys.stderr)
        return 1
    print(f"Tag {argv[1].strip()} matches version.py ({__version__}).")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
