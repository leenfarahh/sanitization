; Inno Setup script for "Sanitize Setup.exe". Build it with: python build_installer.py
; That script stages the app with its own copy of Python and passes the defines below, so compiling this file on its
; own does not work.

#ifndef StageDir
  #error Build the installer with: python build_installer.py
#endif

#define AppName "Sanitize"
#define AppTarget "{app}\python\pythonw.exe"
#define AppAbout "Mask a deck before it goes out, restore it when it comes back"

[Setup]
; Never change the AppId: an installer with the same AppId updates the installed copy instead of adding a second one.
AppId={{1787D575-2B1C-4D8A-9F6B-934C05EB56A6}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=Prezlab
AppComments={#AppAbout}
VersionInfoVersion={#AppVersion}
VersionInfoCompany=Prezlab
VersionInfoDescription={#AppName} Setup
; For the current user, without admin rights. IT can install for everyone with /ALLUSERS on the command line.
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=commandline
DefaultDirName={autopf}\{#AppName}
DisableDirPage=yes
DisableProgramGroupPage=yes
DisableReadyPage=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
OutputBaseFilename={#AppName} Setup
SetupIconFile=sanitize.ico
UninstallDisplayIcon={app}\sanitize.ico
UninstallDisplayName={#AppName}
WizardStyle=modern dynamic
WizardImageFile={#WizardSide}
WizardSmallImageFile={#WizardTop}
Compression=lzma2/max
SolidCompression=yes
CloseApplications=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[InstallDelete]
; An update replaces the whole Python copy, so files from an older Python don't linger next to the new ones.
Type: filesandordirs; Name: "{app}\python"
Type: filesandordirs; Name: "{app}\__pycache__"

[Files]
Source: "{#StageDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; The AppUserModelID matches the one sanitize_desktop.py sets, so the running app shares the shortcut's taskbar button.
Name: "{autoprograms}\{#AppName}"; Filename: "{#AppTarget}"; Parameters: """{app}\sanitize_desktop.py"""; WorkingDir: "{app}"; IconFilename: "{app}\sanitize.ico"; Comment: "{#AppAbout}"; AppUserModelID: "Prezlab.Sanitize"
Name: "{autodesktop}\{#AppName}"; Filename: "{#AppTarget}"; Parameters: """{app}\sanitize_desktop.py"""; WorkingDir: "{app}"; IconFilename: "{app}\sanitize.ico"; Comment: "{#AppAbout}"; AppUserModelID: "Prezlab.Sanitize"; Tasks: desktopicon

[Run]
Filename: "{#AppTarget}"; Parameters: """{app}\sanitize_desktop.py"""; WorkingDir: "{app}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: postinstall nowait skipifsilent

[UninstallDelete]
; Bytecode Python wrote after installing, and any highlighted previews left behind (they hold a deck's original text).
Type: filesandordirs; Name: "{app}\python"
Type: filesandordirs; Name: "{app}\__pycache__"
Type: filesandordirs; Name: "{%TEMP}\Sanitize previews"
