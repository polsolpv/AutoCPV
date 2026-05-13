$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

py -m pip install --upgrade pip
py -m pip install -r requirements.txt
py -m PyInstaller --clean --noconfirm AutoCPV.spec

$isccCandidates = @(
    "$env:ProgramFiles(x86)\Inno Setup 6\ISCC.exe",
    "$env:ProgramFiles\Inno Setup 6\ISCC.exe",
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
)

$iscc = $isccCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) {
    $cmd = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if ($cmd) {
        $iscc = $cmd.Source
    }
}

if (-not $iscc) {
    Write-Warning "Inno Setup no esta instal.lat. S'ha creat dist\AutoCPV.exe, pero no l'instal.lador."
    Write-Warning "Instal.la Inno Setup 6 i torna a executar build-release.ps1."
    exit 0
}

& $iscc installer.iss

Write-Host "Build completada:"
Write-Host " - dist\AutoCPV.exe"
Write-Host " - installer-dist\AutoCPV-Setup.exe"
