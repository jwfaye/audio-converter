@echo off
REM Build script for Audio Converter Windows installer
REM Prerequisites: Python 3.10+, Inno Setup 6

echo === Step 1/2: Building executable with PyInstaller ===
pip install pyinstaller
pyinstaller audio_converter.spec
if errorlevel 1 (
    echo ERROR: PyInstaller build failed.
    exit /b 1
)

echo === Step 2/2: Building installer with Inno Setup ===
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
if errorlevel 1 (
    echo ERROR: Inno Setup build failed.
    exit /b 1
)

echo.
echo === Build complete ===
echo Installer: dist\AudioConverter-Setup.exe
