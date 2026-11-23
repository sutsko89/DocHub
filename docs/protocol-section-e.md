# Протокол замеров раздела Е (baseline «до/после» на целевой Windows)

Назначение: подтвердить на целевой машине Windows эффекты карточек
PERF-2 (2.22.0), STORE-1A (2.24.0), TREE-1 (2.25.0) и контрольные замеры
этапа 5 (START-1, LAZY-1, P-08). Черновые замеры
сняты в Linux-песочнице и в обещания эффекта не подставляются: до
прогона этого протокола соответствующие позиции числятся
«не проверено». Протокол выполняет пользователь на своей Windows-машине;
результаты (stdout прогонов и perf.log) передаются обратно для занесения
в CHANGELOG.

## 1. Что подтверждаем и пороги

| Блок | Карточка | Метрика | Порог |
| --- | --- | --- | --- |
| E-1 | STORE-1A | точечная правка + `save()` @10k (медиана) | ≤ 15 мс |
| E-1 | STORE-1A | `save()` без изменений @10k | ≈ 0 мс (запись не выполняется вовсе) |
| E-1 | STORE-1A | счётчики статусной строки @10k / @50k | сравнение с репликой P-10 (черновик: 1,3 / 4,4 мс против 4,1 / 25,4 мс) |
| E-2 | PERF-2 | p95 dt тика `pointer.poller` при жестах | ≤ 25 мс (предварительный) |
| E-2 | PERF-2 | операции реестра не прерывают dt тики пулера | p95 тиков до и после операций не ухудшается |
| E-3 | TREE-1 | первое отображение списка Treeview @10k | ≤ 300 мс |
| E-3 | TREE-1 | максимальный тик прокрутки Treeview | ≤ 50 мс |
| E-4 | START-1 | максимум duration-записей главного потока в первые 5 с после `startup.built` | ≤ 100 мс |
| E-4 | LAZY-1 | рабочая память процесса при открытом реестре @10k (WorkingSet) | цель −40% к «до» (предварительная) |
| E-4 | P-08 | медиана канала «Запись .docx» (perf.log) | ≤ 50 мс (ориентир) |
| E-4 | P-08 | полный путь «сборка + запись» (bench_docx_build) | ≤ 100 мс (ориентир) |

Состояния кода:

| Состояние | Коммит | Что в нём |
| --- | --- | --- |
| «до» PERF-2 | `961db88` (v2.21.0) | жесты с `update_idletasks`, полная запись базы |
| «до» STORE-1A | `7465dfa` (v2.23.0) | `save()` всегда перезаписывает базу целиком, счётчики по спискам в памяти |
| «до» этапа 5 | `7bea260` (v2.25.2) | сверка/очистки/сроки на UI-потоке (after-цепочка 600/900/1500/1900), тела документов читаются при старте |
| «после» | `main` | инкрементальная запись, Treeview-реестр, стартовый воркер, ленивые тела |

## 2. Подготовка машины (однократно)

1. Windows 10/11 x64. Нужны: **Python 3.13 (64-бит)** — при установке
   отметить «Add python.exe to PATH»; **Git для Windows**.
2. Папка для замеров, например `C:\bench` — вне синхронизации OneDrive/
   Dropbox и вне папки с реальной базой программы.
3. Условия чистоты замера (на каждый сеанс):
   * ноутбук — в розетке; план питания «Высокая производительность»:
     `powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c`;
   * закрыть браузеры, мессенджеры, фоновые загрузки; не трогать мышь и
     не переключать окна во время автоматических прогонов;
   * антивирус не отключать; допустимо добавить каталог `C:\bench` в
     исключения сканирования в реальном времени — этого достаточно;
   * **замеры никогда не идут на реальной базе пользователя**: только
     синтетика `tools/generate_registry.py` и изолированный профиль
     (раздел 6).

## 3. Установка окружения

```powershell
mkdir C:\bench; cd C:\bench
git clone https://github.com/sutsko89/DocHub.git DocHub
cd DocHub
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -c "import storage, perflog; print('ok')"
```

Если PowerShell запрещает `Activate.ps1`: один раз выполнить
`Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` либо вызывать
python напрямую: `.\.venv\Scripts\python.exe` вместо активации.

Зафиксировать метаданные (входят в отчёт):

```powershell
git rev-parse HEAD; python --version
Get-CimInstance Win32_Processor | Select-Object -ExpandProperty Name
Get-CimInstance Win32_ComputerSystem | Select-Object TotalPhysicalMemory
systeminfo | findstr /C:"OS Name" /C:"OS Version"
```

## 4. Общие правила

* Одна машина, одни данные: генератор детерминирован (SEED зашит), базы
  `reg10k`/`reg50k` создаются один раз и используются во всех прогонах
  «до» и «после».
* Первый после запуска ОС прогон — прогрев, в отчёт не входит; далее 5
  повторов, сравниваем медиану и p95.
* Скрипт-замер всегда работает с копией базы во временном каталоге —
  исходники `reg10k`/`reg50k` не изменяются ни на одном шаге.
* «До» и «после» снимать в один день, в одинаковых условиях.

## 5. Блок E-1 — STORE-1A: запись базы

### 5.1 Данные

```powershell
python tools\generate_registry.py --out C:\bench\reg10k --records 10000
python tools\generate_registry.py --out C:\bench\reg50k --records 50000
```

Сохранить вывод обеих команд (строки `save(): N мс` — контрольная точка
генератора).

### 5.2 «После» (текущий main)

```powershell
python tools\bench_store_1a.py --db C:\bench\reg10k\database.db --repeats 5 | Tee-Object C:\bench\e1-after-10k.txt
python tools\bench_store_1a.py --db C:\bench\reg50k\database.db --repeats 5 | Tee-Object C:\bench\e1-after-50k.txt
```

Ожидаемые строки механизма: `Механизм счётчиков: status_counts()…`,
`Механизм save(): инкрементальный + save_full() доступен…`.

### 5.3 «До» (v2.23.0, коммит `7465dfa`)

Отдельное дерево того же репозитория (venv общий, повторно не ставится):

```powershell
cd C:\bench\DocHub
git worktree add C:\bench\dochub-before 7465dfa
copy tools\bench_store_1a.py C:\bench\dochub-before\tools\
cd C:\bench\dochub-before
C:\bench\DocHub\.venv\Scripts\python.exe tools\bench_store_1a.py --db C:\bench\reg10k\database.db --repeats 5 | Tee-Object C:\bench\e1-before-10k.txt
C:\bench\DocHub\.venv\Scripts\python.exe tools\bench_store_1a.py --db C:\bench\reg50k\database.db --repeats 5 | Tee-Object C:\bench\e1-before-50k.txt
```

Ожидаемые строки: `Механизм счётчиков: списки в памяти…`,
`Механизм save(): полная перезапись…`. По завершении:

```powershell
cd C:\bench\DocHub; git worktree remove ..\dochub-before --force
```

### 5.4 Чтение результата

* Критерий аудита: медиана «Точечная правка + save()» @10k ≤ 15 мс.
* `save()` без изменений: «после» ≈ 0 (запись пропускается, ревизия не
  растёт), «до» ≈ стоимости полной перезаписи.
* Полная перезапись `save_full()` печатается в «после» — она же является
  «до»-значением по определению.

## 6. Блок E-2 — PERF-2: жесты окна (перенос, растягивание, snap)

> **Примечание (архитектурное решение от 23.11.2026).** Следующий
> интерфейс DocHub — Python + PySide6 / Qt Widgets со стандартной
> системной рамкой Windows, без самописного переноса и растягивания.
> Блок E-2 измеряет именно самописные жесты текущего Tk-интерфейса:
> он остаётся в итоговой таблице текущей версии, но в пороги после
> перехода на Qt не переносится — пунктов «жесты» в новом коде не будет.

### 6.1 Подготовка сеанса

Замеры жестов идут на копии синтетической базы в отдельном профиле —
реальный `%APPDATA%\DogovorGenerator` не затрагивается:

```powershell
mkdir C:\bench\profile\DogovorGenerator
Copy-Item C:\bench\reg10k\database.db C:\bench\profile\DogovorGenerator\database.db
cd C:\bench\DocHub
$env:APPDATA = "C:\bench\profile"
$env:EWDOCHUB_PERF_LOG = "1"
.\.venv\Scripts\python.exe main.py
```

### 6.2 Сценарий жестов (в обоих состояниях кода одинаковый)

1. Открыть реестр, дождаться списка.
2. **Перенос**: перетащить окно за заголовок плавным круговым движением
   ~10 секунд; повторить 3 раза с паузами 1–2 с.
3. **Растягивание**: тянуть за край/угол ~10 секунд; повторить 3 раза.
4. **Snap**: перетащить окно к краю экрана до появления подложки,
   отпустить; вернуть окно; повторить 3 раза.
5. **Операции реестра** (критерий «не прерывают тики»): сделать точечную
   правку карточки («Сохранить»), удалить строку, нажать «Показать ещё»,
   снова открыть карточку — после этого повторить перенос ~10 с.
6. Закрыть приложение штатно (крестик) — буфер perf-журнала сбрасывается
   при выходе.

`perf.log` окажется в `C:\bench\profile\DogovorGenerator\logs\perf.log`
(та же папка открывается кнопкой «Настройки → Открыть папку журналов»).

### 6.3 «После», отчёт, «до»

```powershell
# после сеанса «после» — открыть новый PowerShell (APPDATA не переопределён)
cd C:\bench\DocHub
.\.venv\Scripts\python.exe tools\bench_store_1a.py --report C:\bench\profile\DogovorGenerator\logs\perf.log | Tee-Object C:\bench\e2-after.txt
# переложить perf.log, чтобы не смешивать прогоны:
Move-Item C:\bench\profile\DogovorGenerator\logs\perf.log C:\bench\perf-after.log
```

Сеанс «до» (v2.21.0, коммит `961db88`): повторить разделы 6.1–6.2,
запустив приложение из старого дерева:

```powershell
cd C:\bench\DocHub
git worktree add C:\bench\dochub-perf 961db88
cd C:\bench\dochub-perf
$env:APPDATA = "C:\bench\profile"
$env:EWDOCHUB_PERF_LOG = "1"
C:\bench\DocHub\.venv\Scripts\python.exe main.py
```

После закрытия приложения скопировать отчётный скрипт в старое дерево и
снять отчёт:

```powershell
copy C:\bench\DocHub\tools\bench_store_1a.py C:\bench\dochub-perf\tools\
C:\bench\DocHub\.venv\Scripts\python.exe tools\bench_store_1a.py --report C:\bench\profile\DogovorGenerator\logs\perf.log | Tee-Object C:\bench\e2-before.txt
```

По завершении: `git worktree remove ..\dochub-perf --force` из
`C:\bench\DocHub`.

* Сравнивать строки `tick pointer.poller` (p95): это dt тиков пулера
  жестов; `registry.canvas` — раскладка списка (Treeview, только «после»);
  `duration storage` — операции записи.
* Критерий: p95 `pointer.poller` ≤ 25 мс в «после» и без деградации после
  операций реестра (шаг 5).

## 7. Блок E-3 — TREE-1: реестр на ttk.Treeview

```powershell
cd C:\bench\DocHub
.\.venv\Scripts\Activate.ps1
python tools\bench_tree_prototype.py --records 1000  --arms all | Tee-Object C:\bench\e3-1k.txt
python tools\bench_tree_prototype.py --records 10000 --arms all | Tee-Object C:\bench\e3-10k.txt
python tools\bench_tree_prototype.py --records 50000 --arms tree | Tee-Object C:\bench\e3-50k.txt
```

Во время прогона на экране мелькает окно — не трогать его мышью.

* Арм `rows` = «до»-механизм (строки из CTk-виджетов, как refresh_full
  2.24.0); на 50k заведомо минуты, поэтому не измеряется.
* Арм `tree` = «после» (ttk.Treeview). Критерии: `populate_ms` @10k
  ≤ 300 мс; `scroll_max_tick_ms` ≤ 50 мс; популяция @50k — для решения
  R-05b о лимите отображения (черновик: 447 мс → мягкий лимит 10 000).

## 8. Блок E-4 — этап 5: старт, память и P-08

Каналы стартовых шагов (`startup.*`) и события `startup.built`/
`startup.ready` появились в 2.28.1; в состоянии «до» этапа 5 (`7bea260`)
их нет. Поэтому строгий критерий ≤ 100 мс проверяется только на «после»
по perf.log; «до»-прогон снимается внешним секундомером (п. 8.4).

### 8.1 Подготовка

Профиль и флаг — как в разделе 6.1 (копия reg10k в изолированный профиль,
`EWDOCHUB_PERF_LOG=1`). Для замера памяти:

```powershell
Get-Process python | Select-Object Id, @{n='WS_MB';e={[math]::Round($_.WorkingSet64/1MB,1)}}
```

### 8.2 Прогон «после» (текущий main)

1. Запустить приложение из `C:\bench\DocHub` (флаг perf, изолированный
   профиль — те же команды, что в 6.1, без worktree).
2. Открыть реестр, дождаться списка, подождать ~30 с (воркер старта
   завершает сверку), сделать 3–5 сохранений договора и спецификации,
   при открытом реестре снять WorkingSet (команда 8.1), закрыть штатно
   (крестик) — буфер perf.log сбрасывается при выходе.
3. Отчёт и ориентир P-08:

```powershell
cd C:\bench\DocHub
.\.venv\Scripts\python.exe tools\bench_store_1a.py --report C:\bench\profile\DogovorGenerator\logs\perf.log | Tee-Object C:\bench\e4-after.txt
.\.venv\Scripts\python.exe tools\bench_docx_build.py --runs 10 --items 30 | Tee-Object C:\bench\e4-p08-after.txt
```

Чтение отчёта:

* строки `event startup.built` и `event startup.ready` — границы окна
  старта (значение 0, метка момента);
* **главная метрика**: все duration-записи с таймстампом в пределах
  5 секунд после `startup.built` на главном потоке — максимум ≤ 100 мс.
  Каналы главного потока: `startup.apply_*`, `startup.backup`,
  `Перерисовка реестра`, `Чтение реестра`, `Запись .docx`. Каналы
  `startup.temp/trash_count/orphans/notice` идут в рабочем потоке и в
  критерий не входят (но именно они показывают, сколько диск ушёл с
  главного потока: в «до» это же время блокировало UI);
* P-08: медиана `Запись .docx` из perf.log (только запись) плюс полный
  путь «сборка + запись» из bench_docx_build.

### 8.3 Память (LAZY-1): «до» и «после»

```powershell
cd C:\bench\DocHub
git worktree add C:\bench\dochub-lazy 7bea260
cd C:\bench\dochub-lazy
$env:APPDATA = "C:\bench\profile"; $env:EWDOCHUB_PERF_LOG = "1"
C:\bench\DocHub\.venv\Scripts\python.exe main.py
# открыть реестр @10k, снять WorkingSet командой 8.1, закрыть
cd C:\bench\DocHub; git worktree remove ..\dochub-lazy --force
```

Шаги «до» и «после» — одинаковые (одна база reg10k, одно время сеанса).
Цель аудита −40% — предварительная: в таблицу заносится фактическое
значение; обещания эффекта не даются до заполненной таблицы.

### 8.4 Старт «до» (внешний секундомер)

На обоих состояниях (main и `7bea260`) трижды зафиксировать: время от
запуска до появления окна и до «отлипа» (список реестра отвечает на
клик). Это сравнение эффекта для итоговой таблицы; строгий критерий
≤ 100 мс — только на «после» по perf.log (п. 8.2).

## 9. Шаблон отчёта

| Блок | Метрика | До | После | Порог | Вердикт |
| --- | --- | --- | --- | --- | --- |
| E-1 @10k | точечная правка + save(), мс (медиана) | | | ≤ 15 | |
| E-1 @10k | save() без изменений, мс | | | ≈ 0 | |
| E-1 @10k | счётчики статусной строки, мс | | | быстрее реплики P-10 | |
| E-1 @50k | точечная правка + save(), мс | | | — | |
| E-2 | p95 dt тика pointer.poller, мс | | | ≤ 25 | |
| E-2 | p95 тиков после операций реестра, мс | | | не хуже | |
| E-3 @10k | populate Treeview, мс | (rows) | | ≤ 300 | |
| E-3 @10k | максимальный тик прокрутки, мс | | | ≤ 50 | |
| E-3 @50k | populate Treeview, мс | | | — | |
| E-4 | старт «до появления окна / до отлипа», с (секундомер, медиана 3 прогонов) | | | — | |
| E-4 | max duration главного потока в первые 5 с после startup.built, мс | — | | ≤ 100 | |
| E-4 | шаги воркера startup.*, мс (медианы) | — | | — | |
| E-4 | WorkingSet @10k, МБ | | | −40% (предв.) | |
| E-4 | «Запись .docx», медиана, мс | | | ≤ 50 (ориентир) | |
| E-4 | bench_docx_build полный путь, мс | | | ≤ 100 (ориентир) | |

К отчёту приложить: stdout всех прогонов (`Tee-Object`-файлы), perf.log
обоих сеансов E-2, метаданные из раздела 3 (коммит, Python, CPU, RAM,
ОС). Если какой-то порог не достигнут или команда завершилась ошибкой —
не править код самостоятельно: прислать исходные файлы как есть.
