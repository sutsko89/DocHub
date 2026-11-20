; Установщик EW.DocHub.
; Компиляция: ISCC.exe installer.iss  (нужен Inno Setup 6/7: https://jrsoftware.org/isinfo.php)
; Результат:  installer\EW.DocHub-<версия>-setup.exe — устанавливает приложение
;             в Program Files (или «только для меня»), создаёт ярлык в меню
;             «Пуск» и (по выбору) на рабочем столе.
;
; Подпись сертификатом (2.16.0):
;   Основной путь — автоматический в CI (.github/workflows/release.yml):
;   добавьте секреты репозитория SIGN_CERT_PFX (base64 .pfx) и
;   SIGN_CERT_PASSWORD — workflow подпишет EW.DocHub.exe и setup.exe через
;   tools/sign_artifact.ps1 (signtool, RFC-3161 таймштамп). Секретов нет —
;   сборка остаётся unsigned. Инструкции по выбору сертификата — в README
;   («Подпись кода»).
;   Ручной путь (локальная сборка с сертификатом):
;   1. Зарегистрируйте SignTool в Inno Setup (Tools → Configure Sign Tools),
;      имя «EWSign», команда — путь к signtool.exe.
;   2. Раскомментируйте строки SignTool/SignedUninstaller в секции [Setup].

#define MyAppName "EW.DocHub"
#define MyAppExeName "EW.DocHub.exe"
; Версия релиза: синхронизируется с APP_VERSION (main.py) скриптом
; tools/bump_version.py. В EWDochub.spec версия из main.py читается напрямую.
#define MyAppVersion "2.19.3"

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
; Версия в имени файла установщика (п. 6.1).
OutputBaseFilename=EW.DocHub-{#MyAppVersion}-setup
SetupIconFile=app_icon.ico
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
; Только Windows 10/11 x64 (п. 6.1).
MinVersion=10.0
ArchitecturesAllowed=x64compatible
; Мьютекс запущенного экземпляра (single_instance.py): Global\EWDocHub_InstallerMutex
; держит каждая копия программы, при отказе в доступе к Global\ — Local\ с тем же
; именем. Установщик не даст обновляться поверх работающего экземпляра (п. 6.1).
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
; Вся папочная сборка PyInstaller целиком
Source: "dist\EW.DocHub\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[InstallDelete]
; Обновление поверх прошлой версии: {app} очищается от файлов старой сборки
; перед копированием новых (состав библиотек PyInstaller onedir может меняться
; между версиями — устаревшие DLL не должны накапливаться в {app}).
; ВАЖНО: пользовательские данные живут в %APPDATA%\DogovorGenerator (реестр,
; настройки, ключ DaData, журнал) и этой секцией НЕ затрагиваются — обновление
; сохраняет все данные пользователя.
Type: filesandordirs; Name: "{app}\*"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Удалить {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; 2.19.2: две записи запуска.
; 1. Чекбокс «Запустить EW.DocHub» на странице «Завершение» — при обычной
;    интерактивной установке (skipifsilent: в тихом режиме финальной
;    страницы нет). Галочка по умолчанию включена, пользователь может снять.
Filename: "{app}\{#MyAppExeName}"; Description: "Запустить {#MyAppName}"; Flags: nowait postinstall skipifsilent
; 2. Тихое обновление (/VERYSILENT, автообновление 2.18.0): программа
;    запускается сама — обновление выглядит как «приложение перезапустилось».
Filename: "{app}\{#MyAppExeName}"; Flags: nowait skipifnotsilent

[Messages]
SetupAppTitle=Установка — {#MyAppName}
SelectDirDesc=Куда установить {#MyAppName}?
ReadyLabel2b=Нажмите «Установить», чтобы начать установку.
FinishedHeadingLabel=Установка {#MyAppName} завершена

[Code]
// П. 6.5: деинсталлятор спрашивает про пользовательские данные и ключ DaData.
// Раньше %APPDATA%\DogovorGenerator (настройки, реестр, журнал, ключ) оставался
// на диске молча. В Pascal-скрипте Inno Setup комментарий — «//», не «;».
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
