; Inno Setup script for Audio Converter
; Requires Inno Setup 6 — https://jrsoftware.org/isdl.php
; Build: "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss

[Setup]
AppName=Audio Converter
AppVersion=0.1.0
AppPublisher=AnotherBrain
DefaultDirName={autopf}\Audio Converter
DefaultGroupName=Audio Converter
UninstallDisplayIcon={app}\audio-converter-gui.exe
OutputDir=dist
OutputBaseFilename=AudioConverter-Setup
Compression=lzma2
SolidCompression=yes
SetupIconFile=assets\app.ico
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible

[Files]
Source: "dist\audio-converter-gui.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "assets\app.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autodesktop}\Audio Converter"; Filename: "{app}\audio-converter-gui.exe"; IconFilename: "{app}\app.ico"
Name: "{group}\Audio Converter"; Filename: "{app}\audio-converter-gui.exe"; IconFilename: "{app}\app.ico"
Name: "{group}\Désinstaller Audio Converter"; Filename: "{uninstallexe}"
