#define MyAppName "AutoCPV"
#define MyAppVersion "1.2"
#define MyAppPublisher "Pol Solsona Franch"
#define MyAppExeName "AutoCPV.exe"
#define MyAppAssocName "AutoCPV"

[Setup]
AppId={{A0A59E52-3A17-4D98-B392-3C86D0E4B67E}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=C:\Users\solso\Documents\New project\installer-dist
OutputBaseFilename=AutoCPV-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile=C:\Users\solso\Documents\New project\assets\logo.ico
WizardImageFile=C:\Users\solso\Documents\New project\assets\installer-wizard.bmp
WizardSmallImageFile=C:\Users\solso\Documents\New project\assets\installer-small.bmp
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "catalan"; MessagesFile: "compiler:Languages\Catalan.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear un accés directe a l'Escriptori"; GroupDescription: "Accessos directes:"

[Files]
Source: "C:\Users\solso\Documents\New project\dist\AutoCPV.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Obrir {#MyAppName}"; Flags: nowait postinstall skipifsilent
