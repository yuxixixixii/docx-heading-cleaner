#define MyAppName "WordNavCleaner"
#define MyAppDisplayName "Word 导航标题清理器"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "WordNavCleaner"
#define MyAppExeName "WordNavCleaner.exe"

[Setup]
AppId={{8D7B8937-0C92-42F6-B860-97DB517E41B4}
AppName={#MyAppDisplayName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppDisplayName}
DisableProgramGroupPage=yes
OutputDir=installer
OutputBaseFilename=WordNavCleanerSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
UninstallDisplayName={#MyAppDisplayName}

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppDisplayName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppDisplayName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppDisplayName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
