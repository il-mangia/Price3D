; Price3D — Inno Setup script (compilare con Inno Setup 6)
; Sorgenti attesi in questa cartella:
;   Price3D_v1.0.exe   (onefile PyInstaller)
;   Price3D.ico        (icona setup e collegamenti)
; Output: Price3D_Setup_v1.0.exe in questa stessa cartella

#define MyAppName "Price3D"
#define MyAppVersion "1.0"
#define MyAppPublisher "il-mangia"
#define MyAppURL "https://github.com/il-mangia/Price3D"
#define MyExeIn "Price3D_v1.0.exe"
#define MyExeOut "Price3D.exe"

[Setup]
AppId={{28da4984-3657-4e43-94dc-13f21540d517}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}/releases/latest
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
LicenseFile=..\LICENSE
SetupIconFile=Price3D.ico
UninstallDisplayIcon={app}\{#MyExeOut}
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes
CloseApplicationsFilter=Price3D*.exe
RestartApplications=no
OutputDir=.
OutputBaseFilename=Price3D_Setup_v1.0
Compression=lzma2/max
SolidCompression=yes
VersionInfoVersion={#MyAppVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName} installer
VersionInfoProductName={#MyAppName}

[Languages]
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "{#MyExeIn}"; DestDir: "{app}"; DestName: "{#MyExeOut}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyExeOut}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyExeOut}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyExeOut}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent
