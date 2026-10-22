; Inno Setup script for Word Document Find & Replace
;
; Installs per-user by default (no administrator rights), because the machines
; this runs on are usually IT-managed and a Program Files install would need an
; admin to approve it. The user can still choose an all-users install if they
; have the rights.
;
; Build:  iscc /DAppVersion=1.0.0 packaging\installer.iss
; Output: dist\installer\DocxFindReplace-Setup-<version>.exe

#define AppName      "Word Document Find & Replace"
#define AppId        "DocxFindReplace"
#define AppExeName   "DocxFindReplace.exe"
#define AppPublisher "Abraham Borg"
#define AppCopyright "Copyright (C) 2025 Abraham Borg. All Rights Reserved."
#define AppURL       "https://github.com/Abe-Borg/docx-find-and-replace"
#define SupportURL   AppURL + "/issues"

; Passed in by the build script; the fallback keeps a bare ISCC run working.
#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif

[Setup]
; Never change AppId - it is how Windows recognises an upgrade of this program
; rather than a second copy of it.
AppId={{A0833429-F7BD-4572-BA5F-9FE53A1E9462}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppCopyright={#AppCopyright}
; Shown in Settings > Apps, so a user or IT has somewhere to ask.
AppPublisherURL={#AppURL}
AppSupportURL={#SupportURL}
AppUpdatesURL={#AppURL}/releases
AppContact={#SupportURL}
VersionInfoVersion={#AppVersion}
VersionInfoCopyright={#AppCopyright}

DefaultDirName={autopf}\{#AppId}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
DisableDirPage=no
AllowNoIcons=yes

; "lowest" installs for the current user with no UAC prompt; the override lets
; someone with admin rights choose an all-users install instead.
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

OutputDir=..\dist\installer
OutputBaseFilename={#AppId}-Setup-{#AppVersion}
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName}

Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
; The interactive wizard shows the license here. /SILENT and /VERYSILENT skip
; every wizard page, this one included, so the application itself asks for
; acceptance on first launch; that is what covers silent and portable use.
LicenseFile=..\LICENSE.txt

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "..\dist\{#AppId}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Also bundled inside the app; installed beside the exe so they can be found.
Source: "..\LICENSE.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "THIRD-PARTY-NOTICES.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Launch {#AppName}"; Flags: nowait postinstall skipifsilent
