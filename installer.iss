#define MyAppName "AutoCPV"
#define MyAppVersion "1.5"
#define MyAppPublisher "AutoCPV"
#define MyAppExeName "AutoCPV.exe"

[Setup]
AppId={{7E843B90-176F-4C05-8D83-A04B49E5D3F8}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=installer-dist
OutputBaseFilename=AutoCPV-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile=assets\logo.ico
WizardImageFile=assets\installer-wizard.bmp
WizardSmallImageFile=assets\installer-small.bmp
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear acceso directo en el escritorio"; GroupDescription: "Accesos directos:"; Flags: checkedonce

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "PROMPT AutoCPV.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "assets\logo-trimmed.png"; DestDir: "{app}\assets"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -Command ""$ErrorActionPreference='SilentlyContinue'; $pf=[Environment]::GetFolderPath('ProgramFiles'); $pf86=[Environment]::GetFolderPath('ProgramFilesX86'); $x64=Join-Path $pf 'PDF24\pdf24-Ocr.exe'; $x86=Join-Path $pf86 'PDF24\pdf24-Ocr.exe'; if (-not (Test-Path $x64) -and -not (Test-Path $x86)) {{ if (Get-Command winget -ErrorAction SilentlyContinue) {{ winget install --id geeksoftwareGmbH.PDF24Creator --source winget --accept-package-agreements --accept-source-agreements }} else {{ Start-Process 'https://tools.pdf24.org/es/creator' }} }}"""; StatusMsg: "Comprobando PDF24 Creator..."; Flags: waituntilterminated; Check: ShouldInstallPDF24
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
function IsPDF24Installed: Boolean;
begin
  Result :=
    FileExists(ExpandConstant('{pf}\PDF24\pdf24-Ocr.exe')) or
    FileExists(ExpandConstant('{pf32}\PDF24\pdf24-Ocr.exe'));
end;

function ShouldInstallPDF24: Boolean;
begin
  Result := not IsPDF24Installed;
end;
