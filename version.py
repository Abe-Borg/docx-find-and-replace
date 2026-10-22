"""
Single source of truth for the application version, name and attribution.

Read by the GUI (window title, About box), the PyInstaller specs, the Windows
file-version resource, and the Inno Setup installer. Bump it here and every
artifact follows. LICENSE.txt and installer.iss repeat the copyright line and
the URLs; tests/test_packaging.py fails if they drift from these.
"""

__version__ = "1.1.0"

APP_NAME = "Word Document Find & Replace"
APP_ID = "DocxFindReplace"          # executable and install-folder name
PUBLISHER = "Abraham Borg"
COPYRIGHT = f"Copyright (C) 2025 {PUBLISHER}. All Rights Reserved."

HOMEPAGE_URL = "https://github.com/Abe-Borg/docx-find-and-replace"
RELEASES_URL = HOMEPAGE_URL + "/releases"
# Where users report problems and ask for permissions the license does not grant.
SUPPORT_URL = HOMEPAGE_URL + "/issues"
