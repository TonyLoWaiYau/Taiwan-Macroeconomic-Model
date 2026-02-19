#define MyAppName "Taiwan Macroeconomic Model v3.5"
#define MyAppVersion "3.5.1"
#define MyAppPublisher "Tony Lo/ Feb 2026"
#define MyAppExeName "run_app.vbs"

[Setup]
AppId={{A2D57C4A-6F2F-4D8E-9B2C-111111111111}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=.
OutputBaseFilename={#MyAppName}_Setup_{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
; Copy everything in your app folder EXCEPT installer folder itself.
Source: "..\app\*";   DestDir: "{app}\app";   Flags: recursesubdirs createallsubdirs
Source: "..\tools\*"; DestDir: "{app}\tools"; Flags: recursesubdirs createallsubdirs
Source: "..\R-4.5.2\*";     DestDir: "{app}\R-4.5.2";     Flags: recursesubdirs createallsubdirs
; Optional template
; Source: "..\data_template\*"; DestDir: "{app}\data_template"; Flags: recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\tools\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\tools\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\tools\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent