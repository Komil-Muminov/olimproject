; Установщик «Распределение по ГИБ» (Inno Setup 6).
; Сначала собрать dist\GIB.exe (build.bat), затем: ISCC installer.iss -> dist\GIB-Setup.exe

#define AppName "Распределение по ГИБ"
#define AppVersion "1.0"

[Setup]
AppId={{6B0F3C1E-8D2A-4F5B-9C71-2E4A9D3B7F10}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=Olim
DefaultDirName={localappdata}\Programs\GIB
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=dist
OutputBaseFilename=GIB-Setup
SetupIconFile=assets\icon.ico
UninstallDisplayIcon={app}\GIB.exe
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "ru"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на рабочем столе"; GroupDescription: "Ярлыки:"

[Files]
Source: "dist\GIB.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\GIB.exe"
Name: "{group}\Удалить {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\GIB.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\GIB.exe"; Description: "Запустить программу"; Flags: nowait postinstall skipifsilent
