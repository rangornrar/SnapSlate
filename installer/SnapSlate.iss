#define MyAppName "SnapSlate"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "SnapSlate"
#define MyAppSetupName "SnapSlate-Setup"

[Setup]
AppId={{DF0C685B-1C0B-4B8A-AE6C-4F0BC56F6D48}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=admin
UsePreviousPrivileges=no
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
Uninstallable=no
SetupIconFile=..\Assets\AppIcon.ico
OutputDir=release
OutputBaseFilename={#MyAppSetupName}
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "french"; MessagesFile: "compiler:Languages\French.isl"

[Files]
Source: "dist\*"; DestDir: "{tmp}\SnapSlate"; Flags: ignoreversion recursesubdirs createallsubdirs deleteafterinstall

[Run]
Filename: "{cmd}"; Parameters: "/c ""{tmp}\SnapSlate\Install-SnapSlate.cmd"""; Flags: waituntilterminated
