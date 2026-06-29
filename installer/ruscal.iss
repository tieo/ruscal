; Inno Setup script for ruscal.
;
; User-scope install (no admin prompt). Places ruscal.exe under
; %LOCALAPPDATA%\Programs\ruscal, creates a Start-menu shortcut,
; registers an Apps & Features entry, and ships a standard uninstaller.
; Autostart is handled in-app (Settings → "Start with Windows" toggle),
; not by the installer, so the persistence registry write only happens
; after explicit user consent.

#define MyAppName    "ruscal"
#define MyAppPublisher "tieo"
#define MyAppURL     "https://github.com/tieo/ruscal"
#define MyAppExeName "ruscal.exe"

; The CI passes -DMyAppVersion=1.2.1 on the iscc.exe command line.
#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif

[Setup]
; AppId is a stable GUID identifying ruscal in Apps & Features.
; Don't change this — collisions across versions are intentional so
; one install upgrades to another instead of accumulating entries.
AppId={{8F8C9F4A-3D71-4E2B-9C2C-1B9D6C2B7E33}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}/issues
AppUpdatesURL={#MyAppURL}/releases
; User-scope install — no admin elevation needed.
PrivilegesRequired=lowest
DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
LicenseFile=..\LICENSE
OutputBaseFilename=ruscal-setup
; Setup icon is intentionally absent — Inno's default wizard glyph is fine.
; The *installed* exe has its proper icon baked in via `winresource` in
; build.rs, so Start-menu and Apps-list entries use the right icon.
Compression=lzma
SolidCompression=yes
WizardStyle=modern
; Show ruscal in Apps & Features.
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "..\target\x86_64-pc-windows-msvc\release\ruscal.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{userprograms}\{#MyAppName}";  Filename: "{app}\{#MyAppExeName}"
Name: "{userdesktop}\{#MyAppName}";   Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Offer to launch ruscal after the installer finishes. nowait so the
; installer can return immediately; skipifsilent so silent installs
; (winget) don't auto-launch a GUI window the user didn't ask for.
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent

; No [UninstallDelete] section: the uninstaller leaves user data
; (%LOCALAPPDATA%\ruscal — tokens, config, state.json) on disk so an
; upgrade-by-reinstall doesn't make the user re-auth Google + re-pair
; calendars. Standard "uninstall doesn't wipe your data" behaviour.
; A future "wipe user data" option can be added as an opt-in task.
