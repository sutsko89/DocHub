; Установщик EW.DocHub.
; Компиляция: ISCC.exe installer.iss  (нужен Inno Setup 6: https://jrsoftware.org/isinfo.php)
; Результат:  installer\install.exe — устанавливает приложение в Program Files,
;             создаёт ярлык в меню «Пуск» и (по выбору) на рабочем столе.

#define MyAppName "EW.DocHub"
#define MyAppExeName "EW.DocHub.exe"
#define MyAppVersion "1.4.0"

[Setup]
AppId={{7B4E6F2A-91C4-4A3E-9B5D-2C8A1D6E4F10}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher=EW.DocHub
DefaultDirName={autopf}\EW.DocHub
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
UninstallDisplayIcon={app}\{#MyAppExeName}
OutputDir=installer
OutputBaseFilename=install
SetupIconFile=app_icon.ico
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на &рабочем столе"; GroupDescription: "Дополнительно:"

[Files]
; Вся папочная сборка PyInstaller целиком
Source: "dist\EW.DocHub\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Удалить {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Запустить {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Messages]
SetupAppTitle=Установка — {#MyAppName}
SelectDirDesc=Куда установить {#MyAppName}?
ReadyLabel2b=Нажмите «Установить», чтобы начать установку.
FinishedHeadingLabel=Установка {#MyAppName} завершена
