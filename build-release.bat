@echo off
setlocal
cd /d "%~dp0"

py -m PyInstaller --noconfirm AutoCPV.spec
if errorlevel 1 exit /b 1

"C:\Users\solso\AppData\Local\Programs\Inno Setup 6\ISCC.exe" "C:\Users\solso\Documents\New project\installer.iss"
if errorlevel 1 exit /b 1

echo.
echo Release compilada:
echo %cd%\dist\AutoCPV.exe
echo %cd%\installer-dist\AutoCPV-Setup.exe
