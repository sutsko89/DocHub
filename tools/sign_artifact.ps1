# Подпись артефакта signtool'ом по секретам CI (2.16.0).
#
# Вызывается из .github/workflows/release.yml только когда заданы секреты
# репозитория SIGN_CERT_PFX (base64 сертификата .pfx) и SIGN_CERT_PASSWORD.
# Секретов нет — шаг пропускается, сборка остаётся рабочей (unsigned).
#
# Использование: ./tools/sign_artifact.ps1 <путь к exe>
param(
    [Parameter(Mandatory = $true)]
    [string]$Path
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $Path)) { throw "Файл не найден: $Path" }
if (-not $env:SIGN_CERT_PFX) { throw "SIGN_CERT_PFX не задан" }
if (-not $env:SIGN_CERT_PASSWORD) { throw "SIGN_CERT_PASSWORD не задан" }

# PFX из секрета: base64, допускаются переводы строк — вычищаем пробелы
# и разворачиваем во временный файл (RUNNER_TEMP живёт только в рамках job).
$pfx = Join-Path $env:RUNNER_TEMP "sign_cert.pfx"
$b64 = ($env:SIGN_CERT_PFX -replace '\s', '')
[IO.File]::WriteAllBytes($pfx, [Convert]::FromBase64String($b64))

try {
    # Свежий signtool из Windows SDK (предустановлен на windows-latest);
    # выбираем самый новый каталог версий, только 64-битный бинарник.
    $signtool = Get-ChildItem "C:\Program Files (x86)\Windows Kits\10\bin" -Recurse -Filter signtool.exe |
        Where-Object { $_.FullName -match '\\x64\\' } |
        Sort-Object FullName -Descending |
        Select-Object -First 1
    if (-not $signtool) { throw "signtool.exe не найден в Windows SDK" }

    # RFC-3161 таймштамп: подпись остаётся валидной и после истечения
    # срока сертификата (/td SHA256 вместе с /fd SHA256).
    & $signtool.FullName sign /fd SHA256 /td SHA256 /tr http://timestamp.digicert.com /f $pfx /p $env:SIGN_CERT_PASSWORD $Path
    if ($LASTEXITCODE -ne 0) { throw "signtool sign завершился с ошибкой $LASTEXITCODE" }

    # Контроль: подпись должна проходить проверку Authenticode.
    & $signtool.FullName verify /pa $Path
    if ($LASTEXITCODE -ne 0) { throw "Проверка подписи не удалась ($LASTEXITCODE)" }
    Write-Host "Подписан и проверен: $Path"
}
finally {
    Remove-Item $pfx -ErrorAction SilentlyContinue
}
