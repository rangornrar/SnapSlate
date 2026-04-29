#define MyAppName "SnapSlate"
#define MyAppVersion "2026.4.26.1"
#define MyAppPublisher "SnapSlate"
#define MyAppSetupName "SnapSlate-Setup"

[Setup]
AppId={{DF0C685B-1C0B-4B8A-AE6C-4F0BC56F6D48}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=admin
UsePreviousPrivileges=no
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
Uninstallable=yes
SetupIconFile=..\Assets\AppIcon.ico
OutputDir=release
OutputBaseFilename={#MyAppSetupName}
ArchitecturesInstallIn64BitMode=x64compatible
ShowLanguageDialog=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "french"; MessagesFile: "compiler:Languages\French.isl"

[Tasks]
Name: "desktopicon"; Description: "&Create a desktop icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "dist\app\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "dist\README.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\SnapSlate"; Filename: "{app}\SnapSlate.exe"
Name: "{autodesktop}\SnapSlate"; Filename: "{app}\SnapSlate.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\SnapSlate.exe"; Description: "Launch SnapSlate"; Flags: nowait postinstall skipifsilent
