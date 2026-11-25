; Установщик EW.DocHub — Qt-сборка (этап M7 миграции Tk → PySide6).
; Компиляция: ISCC.exe installer_qt.iss  (нужен Inno Setup 6/7: https://jrsoftware.org/isinfo.php)
; Результат:  installer\EW.DocHub-<версия>-qt-setup.exe — устанавливает
;             приложение в Program Files (или «только для меня»), создаёт
;             ярлык в меню «Пуск» и (по выбору) на рабочем столе.
;
; КЛЮЧЕВОЕ ОТЛИЧИЕ ОТ installer.iss: папочная сборка PyInstaller —
; dist\EW.DocHub-qt\ (main_qt.py, PySide6), имя exe — EW.DocHub-qt.exe,
; имя ассета — EW.DocHub-<версия>-qt-setup.exe (суффикс флейвора, И.2).
;
; ВАЖНО (критерий CP2): AppId и AppMutex — ТЕ ЖЕ, что у Tk-сборки:
;   * установка Qt поверх Tk 2.28.2 работает как обновление — та же папка
;     {autopf}\EW.DocHub, ярлыки переименовываются на тот же EW.DocHub;
;   * мьютексы single_instance не дают обновляться поверх работающего
;     экземпляра ЛЮБОЙ сборки (Tk или Qt);
;   * пользовательские данные (%APPDATA%\DogovorGenerator: реестр,
;     настройки, ключ DaData, журнал) не затрагиваются.
;
; Подпись сертификатом — как у Tk-сборки: секреты SIGN_CERT_PFX /
; SIGN_CERT_PASSWORD в CI (tools/sign_artifact.ps1) или ручной SignTool
; «EWSign» (см. installer.iss и README, «Подпись кода»).

#define MyAppName "EW.DocHub"
#define MyAppExeName "EW.DocHub-qt.exe"
; Версия релиза: синхронизируется с APP_VERSION (app_common.py) скриптом
; tools/bump_version.py (M7: скрипт обновляет оба установщика).
#define MyAppVersion "2.28.2"

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
; Версия в имени файла установщика + суффикс флейвора (И.2 плана миграции).
OutputBaseFilename=EW.DocHub-{#MyAppVersion}-qt-setup
SetupIconFile=app_icon.ico
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
; Только Windows 10/11 x64 (п. 6.1).
MinVersion=10.0
ArchitecturesAllowed=x64compatible
; Мьютекс запущенного экземпляра (single_instance.py): Global\EWDocHub_InstallerMutex
; держит каждая копия программы (Tk и Qt), при отказе в доступе к Global\ —
; Local\ с тем же именем. Установщик не даст обновляться поверх работающего
; экземпляра (п. 6.1); Qt-установщик конфликтует и с Tk, и с Qt — намеренно.
AppMutex=Global\EWDocHub_InstallerMutex;Local\EWDocHub_InstallerMutex
; Файлы заняты запущенной программой — предлагаем закрыть её перед установкой.
CloseApplications=yes
; Установка «только для меня» без прав администратора — по выбору в диалоге (п. 6.1).
PrivilegesRequiredOverridesAllowed=dialog
; Страница лицензии и уведомлений о сторонних библиотеках (п. 6.1).
LicenseFile=license.txt
; Подпись (см. шапку файла): раскомментировать при наличии сертификата и
; зарегистрированного SignTool-инструмента «EWSign».
; SignTool=EWSign $f
; SignedUninstaller=yes

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на &рабочем столе"; GroupDescription: "Дополнительно:"

[Files]
; Вся папочная сборка PyInstaller Qt-версии целиком
Source: "dist\EW.DocHub-qt\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[InstallDelete]
; Обновление поверх прошлой версии (Tk 2.28.2 или прошлый Qt): {app}
; очищается от файлов старой сборки перед копированием новых — состав
; библиотек PyInstaller onedir меняется между версиями и сборками (стек
; Tk уходит, стек Qt приходит), устаревшие DLL не должны накапливаться.
; ВАЖНО: пользовательские данные живут в %APPDATA%\DogovorGenerator (реестр,
; настройки, ключ DaData, журнал) и этой секцией НЕ затрагиваются —
; обновление Tk → Qt сохраняет все данные пользователя.
Type: filesandordirs; Name: "{app}\*"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Удалить {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Две записи запуска — как у Tk-сборки (installer.iss, 2.19.2):
; 1. Чекбокс «Запустить EW.DocHub» на странице «Завершение» (skipifsilent).
; 2. Тихое обновление (/VERYSILENT, автообновление): программа запускается
;    сама — обновление выглядит как «приложение перезапустилось».
Filename: "{app}\{#MyAppExeName}"; Description: "Запустить {#MyAppName}"; Flags: nowait postinstall skipifsilent
Filename: "{app}\{#MyAppExeName}"; Flags: nowait skipifnotsilent

[Messages]
SetupAppTitle=Установка — {#MyAppName}
SelectDirDesc=Куда установить {#MyAppName}?
ReadyLabel2b=Нажмите «Установить», чтобы начать установку.
FinishedHeadingLabel=Установка {#MyAppName} завершена

[Code]
// П. 6.5: деинсталлятор спрашивает про пользовательские данные и ключ DaData
// (тот же сценарий, что у Tk-сборки). Комментарий в Inno Setup — «//».
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
    if MsgBox(
        'Удалить настройки, реестр документов, журнал и ключ DaData?' + #13#10 +
        '(%APPDATA%\DogovorGenerator)' + #13#10#13#10 +
        'Папка документов, если вы перенесли её в другое место (например, в ' +
        '«Документы»), удалена не будет — при желании удалите её вручную.',
        mbConfirmation, MB_YESNO) = IDYES then
      DelTree(ExpandConstant('{userappdata}\DogovorGenerator'), True, True, True);
end;
