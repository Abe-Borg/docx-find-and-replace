"""
Single source of truth for the application version and name.

Read by the GUI (window title), the PyInstaller specs, the Windows file-version
resource, and the Inno Setup installer. Bump it here and every artifact follows.
"""

__version__ = "1.1.0"

APP_NAME = "Word Document Find & Replace"
APP_ID = "DocxFindReplace"          # executable and install-folder name
PUBLISHER = "Abraham Borg"
