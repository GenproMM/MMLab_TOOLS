# AGENTS.md — стандарт скриптов MM LAB

**Назначение.** Канонический стандарт написания pyRevit-скриптов MM LAB: правила кода,
структура кнопки, совместимость версий Revit, машинная проверка конвенции, каталог команд
и Git-регламент.

**Аудитория.** Новички без опыта Python и Revit API — и ИИ-агенты: Claude Code, Gemini CLI,
Kilo Code. Текст стандарта — на русском; код, имена Revit API и имена файлов — английские
или как в репозитории.

**Статус.** Единственный источник правды по конвенции. Per-agent файлы (`CLAUDE.md`,
`GEMINI.md`, `.kilocode/rules/00-mmlab.md`) — тонкие указатели на этот файл плюс специфика
конкретного агента; текст конвенции в них НЕ дублируется.

## Архитектура репозитория

```text
MMLab_TOOLS/
├── AGENTS.md                        # этот стандарт (единственный источник правды)
├── CLAUDE.md / GEMINI.md            # тонкие указатели агентов (@AGENTS.md)
├── .kilocode/rules/00-mmlab.md      # указатель для старых версий Kilo Code
├── MM LAB.extension/                # расширение pyRevit
│   ├── lib/                         # first-party библиотека: revit_compat.py и хелперы
│   └── MM Lab.tab/                  # вкладка MM Lab
│       ├── bundle.yaml              # layout вкладки (порядок панелей)
│       └── <Панель>.panel/          # панель (АРХИТЕКТУРА / ИОС / КООРДИНАЦИЯ)
│           ├── bundle.yaml          # layout панели (порядок и видимость кнопок)
│           └── <Кнопка>.pushbutton/ # кнопка: script.py + bundle.yaml + README.md + icon.png
├── lib/                             # vendored-библиотеки: openpyxl, et_xmlfile
├── templates/                       # шаблон новой кнопки (НоваяКнопка.pushbutton)
├── tools/                           # check_convention.py, convention_baseline.json, tests/
└── agents/commands/                 # канонические процедуры mm-команд
```

- **First-party lib** — `MM LAB.extension/lib`: pyRevit сам добавляет этот каталог в
  `sys.path` всем командам расширения. Правки в `lib` подхватываются только после
  pyRevit Reload (или перезапуска Revit).
- **Vendored lib** — корневой `lib/` (openpyxl, et_xmlfile): в `sys.path` автоматически
  НЕ попадает; подключается только вызовом `revit_compat.ensure_vendor_lib()` и только
  теми кнопками, которым vendored-пакет действительно нужен.
- **templates/** лежит вне `MM Lab.tab`, поэтому pyRevit не загружает шаблон как кнопку.
- **tools/** — инструменты проверки, запускаются обычным CPython без Revit.
- **agents/commands/** — канонические процедуры команд `/mm-*` (по одному файлу на команду).

## Обязательные правила кода

1. **Шапка файла** [MM001, MM002, MM004]. Строка 1 — `#! python3`, строка 2 —
   `# -*- coding: utf-8 -*-`, затем русский docstring со строками
   «Совместимость: Revit 2020 / 2022 / 2024» и «Зависимости:», затем `__title__`
   и `__author__`. Точный блок:

   ```python
   #! python3
   # -*- coding: utf-8 -*-
   """Название кнопки

   Что делает кнопка, кратко.

   Совместимость: Revit 2020 / 2022 / 2024
   Зависимости: нет
   """

   __title__ = "Новая\nкнопка"
   __author__ = "GENPRO LAB"
   ```

2. **Кодировка UTF-8 без BOM** [MM003]. Байты `EF BB BF` в начале файла запрещены:
   BOM перед шебангом ломает выбор движка. В редакторе выбирай «UTF-8», а не «UTF-8 with BOM».
3. **Только CPython.** IronPython для нового кода запрещён (шебанг `#! python3` обязателен).
   Синтаксис — консервативный: без `match` и walrus-оператора `:=` — движок pyRevit может
   быть старше локального Python. Чекер — статический гейт; runtime-истина — проверка в Revit (UAT).
4. **Импорты только из белого списка** [MM008]: stdlib + хост-платформа (`clr`, `System`,
   `Autodesk`, `pyrevit`, `Microsoft`) + first-party `MM LAB.extension/lib` + vendored
   (`openpyxl`, `et_xmlfile`). «Без сторонних импортов» означает: pip-пакеты и любые другие
   библиотеки запрещены; ЯВНОЕ исключение — vendored-библиотеки в корневом `lib/`,
   подключаемые через `revit_compat.ensure_vendor_lib()`.
5. **Запрет wildcard-импортов** [MM009]: `from X import *` не используется — импортируй
   имена явно.
6. **Общие функции — в lib.** Логика, нужная двум и более кнопкам, живёт в
   `MM LAB.extension/lib`; дублирование её копий в скриптах запрещено.
7. **Параметры элементов — через revit_compat** [MM010]: `revit_compat.get_parameter(...)`
   с `BuiltInParameter` или `revit_compat.get_shared_parameter(...)` с GUID.
   `LookupParameter("строковый литерал")` в скриптах запрещён — строковые имена
   локализуются и молча ломаются.
8. **Мультиверсия — ТОЛЬКО через revit_compat** (D-01). Ветвления вида `if version >= ...`
   в скриптах кнопок запрещены; единственное место версионных веток — `revit_compat.py`.
9. **Fail-fast по версии Revit** (D-03): `main()` начинается с
   `revit_compat.require_supported_version(COMMAND_NAME)` — на неподдерживаемой версии
   пользователь видит TaskDialog с перечнем поддерживаемых версий, скрипт мягко завершается.
10. **Транзакции**: `transaction.Start()` — ПЕРЕД `try`; `transaction.Commit()` — в `try`;
    в `except` — `transaction.RollBack()` и `raise` (каркас шаблона
    `templates/НоваяКнопка.pushbutton/script.py`). `Start()` внутри `try` запрещён:
    если сам `Start()` упадёт, `except` вызовет `RollBack()` у незапущенной транзакции
    и замаскирует исходную ошибку. Модель меняется только внутри транзакции.
11. **Верхнеуровневая обработка ошибок** [MM011]: вызов `main()` обёрнут в
    `try/except Exception` с показом ошибки через `TaskDialog`. Голый `except:` запрещён —
    всегда указывай класс исключения.
12. **pyrevit.forms под CPython не работает** [MM012]: для диалогов используй `TaskDialog`
    или WinForms (`System.Windows.Forms`).
13. **README.md обязателен** [MM006]: у каждой кнопки — README с назначением и порядком работы.
14. **bundle.yaml обязателен** [MM005]: у каждой кнопки — bundle.yaml с ключами `title:`
    и `tooltip:`.
15. **Регистрация в layout обязательна** [MM007]: кнопка добавляется в `layout:`
    родительского panel-`bundle.yaml`, иначе на pyRevit ≤ 5.x она не видна.
16. **Без мусора в папке кнопки** [MM013]: `__pycache__/`, `*.pyc`, `*.csv`, `.vs/`
    в папке кнопки запрещены.
17. **Канонический lib-бутстрап** [MM014] (D-15). Подключение first-party `lib` — только
    блоком ниже, дословно (он же в шаблоне `templates/НоваяКнопка.pushbutton`).

```python
import os
import sys

_SCRIPT_DIR = os.path.dirname(__file__)
# pushbutton -> panel -> tab -> MM LAB.extension
_EXTENSION_DIR = os.path.normpath(os.path.join(_SCRIPT_DIR, "..", "..", ".."))
_LIB_DIR = os.path.join(_EXTENSION_DIR, "lib")
if os.path.isdir(_LIB_DIR) and _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)
```

Любые другие формы запрещены: легаси-имя переменной «корня расширения» из старых
скриптов (на деле указывала на корень репозитория), подъём `os.path.join` на 4 уровня
`..` и любой `sys.path`-вызов, отличный от `sys.path.insert(0, _LIB_DIR)`.
18. **doc/uidoc — параметрами.** `__revit__`, `doc`, `uidoc` получай в `script.py`
    и передавай в функции аргументами; тянуть `__revit__` из lib-модулей запрещено
    (инжект в lib — негарантированный контракт pyRevit).

## Структура кнопки

Состав папки `<Кнопка>.pushbutton/`:

| Файл | Обязателен | Назначение |
| ------------- | ---------- | ----------------------------------------------- |
| `script.py` | да | код кнопки (шапка → бутстрап → `main()`) |
| `bundle.yaml` | да | подпись и подсказка кнопки |
| `README.md` | да | назначение, порядок работы, ограничения |
| `icon.png` | желателен | иконка ~96×96 PNG |

Формат `bundle.yaml` кнопки (локализованные `ru`/`en_us`, `\n` переносит подпись):

```yaml
title:
  en_us: "New\nButton"
  ru: "Новая\nкнопка"
tooltip:
  en_us: "What this button does, in one sentence"
  ru: "Что делает кнопка, одним предложением"
author: "GENPRO LAB"
```

Правила `layout:` в panel-`bundle.yaml`:

- записи — имена папок кнопок БЕЗ суффикса `.pushbutton`;
- порядок записей = порядок кнопок на панели;
- запись из 3+ символов `-` (`---`) — вертикальный разделитель;
- запись из 3+ символов `>` (`>>>`) — всё ниже уходит в slide-out;
- кнопка, не указанная в `layout:`, на pyRevit ≤ 5.x НЕ отображается;
- каждая запись layout обязана иметь папку на диске (иначе орфан, MM007).

После добавления кнопки или правки `lib` — **pyRevit Reload обязателен** (расширение
парсится на старте Revit).

## Совместимость Revit (2020 / 2022 / 2024)

Поддерживаемые версии — `revit_compat.SUPPORTED_VERSIONS = (2020, 2022, 2024)`.
На неподдерживаемой или неопределённой версии кнопка обязана завершаться fail-fast
с перечнем поддерживаемых версий (правило 9). Ломающие изменения API закрыты хелперами
`MM LAB.extension/lib/revit_compat.py`:

| Область | Было (Revit 2020) | Стало (2022 / 2024) | Хелпер revit_compat |
| ---------- | -------------------------------- | ---------------------------------------------- | ------------------------------------- |
| Units | `DisplayUnitType` (deprecated 2021, удалён 2022) | `ForgeTypeId` / `UnitTypeId` | `convert_from_internal` / `convert_to_internal` |
| Перекрытия | `doc.Create.NewFloor(CurveArray, ...)` (удалён 2023) | `Floor.Create(doc, IList[CurveLoop], ...)` (с 2022) | `create_floor` |
| ElementId | `.IntegerValue` (Int32) | `.Value` (Int64) и `ElementId(Int64)` (с 2024) | `element_id_value` / `make_element_id` |
| pythonnet 3.x | мягкие касты IronPython | строгие overload/enum-касты: `get_Parameter(bip)` падает с TypeError, `int → enum` запрещён, питоновский list не маршалится в `IList[T]` | `get_parameter` / `enum_from_int` / `to_net_list` / `iter_count` |

Публичный API `revit_compat` (скрипты зовут хелперы, а не сырой версионный API):

- `SUPPORTED_VERSIONS` — кортеж поддерживаемых версий Revit: `(2020, 2022, 2024)`.
- `require_supported_version(command_name)` — fail-fast гейт версии в начале `main()`:
  TaskDialog + SystemExit на неподдерживаемой версии; возвращает версию (int).
- `get_parameter(element, built_in_parameter, *fallback_names)` — чтение параметра
  по `BuiltInParameter` каноническим каскадом pythonnet-обходов.
- `get_shared_parameter(element, guid)` — чтение общего параметра по GUID (str или `System.Guid`).
- `element_id_value(element_id)` — числовое значение ElementId: `.Value` (2024+) → `.IntegerValue`.
- `make_element_id(id_value)` — ElementId из числа с учётом Int64 (2024+).
- `convert_from_internal(value, unit_key)` — из внутренних единиц Revit; ключи `"mm" / "cm" / "m" / "m2" / "m3"`.
- `convert_to_internal(value, unit_key)` — во внутренние единицы Revit (те же ключи).
- `create_floor(doc, curve_loops, floor_type_id, level_id)` — перекрытие на всех
  поддерживаемых версиях; транзакцию открывает вызывающий скрипт.
- `to_net_list(items, net_type)` — `System.Collections.Generic.List[T]` из питоновской последовательности.
- `enum_from_int(enum_type, int_value)` — явный каст int → .NET enum через `Enum.ToObject`.
- `iter_count(sequence)` — безопасный подсчёт элементов .NET/py-последовательности (`len` → `.Count` → итерация).
- `ensure_vendor_lib()` — подключает корневой vendored-каталог `lib/` к `sys.path`
  (только когда кнопке нужен openpyxl/et_xmlfile).

## Проверка конвенции

Машинный гейт — `tools/check_convention.py` (обычный CPython ≥ 3.10, без Revit; код кнопок
не исполняется — только чтение и `ast.parse`):

```bash
# одна кнопка (папка *.pushbutton)
py -3 tools/check_convention.py "MM LAB.extension/MM Lab.tab/<Панель>.panel/<Кнопка>.pushbutton"

# сырой .py (правила уровня файла + AST; структурные пропускаются)
py -3 tools/check_convention.py "путь/к/скрипту.py"

# весь репозиторий с baseline легаси-кнопок
py -3 tools/check_convention.py --all --baseline tools/convention_baseline.json

# гейт приёмки новых/адаптируемых кнопок: baseline игнорируется, warning = error
py -3 tools/check_convention.py "<путь>" --strict
```

Exit-коды: 0 — чисто, 1 — есть нарушения (в `--strict` — и warning), 2 — ошибка использования.
Машинный вывод — флаг `--json`.

Коды правил (severity — как в словаре RULES чекера):

| Код | Severity | Проверка |
| ----- | ------- | ------------------------------------------------------------------ |
| MM000 | error | файл не читается или не парсится `ast.parse` |
| MM001 | error | строка 1 script.py — `#! python3` |
| MM002 | error | строка 2 — `# -*- coding: utf-8 -*-` |
| MM003 | error | файл начинается с UTF-8 BOM (EF BB BF) — BOM запрещён |
| MM004 | warning | docstring модуля содержит «Совместимость:» и «Зависимости:» |
| MM005 | error | в папке кнопки есть bundle.yaml с ключами `title:` и `tooltip:` |
| MM006 | error | в папке кнопки есть README.md |
| MM007 | error | кнопка/запись не согласована с layout bundle.yaml (регистрация + орфаны) |
| MM008 | error | импорт вне белого списка (stdlib, clr/System/Autodesk/pyrevit, lib, vendored) |
| MM009 | error | wildcard-импорт `from X import *` запрещён |
| MM010 | warning | `LookupParameter("строковый литерал")` — используй revit_compat |
| MM011 | warning | голый `except:` — укажи класс исключения |
| MM012 | warning | импорт pyrevit.forms (не работает под CPython3) |
| MM013 | warning | мусор в папке кнопки (`__pycache__/`, `*.pyc`, `.vs/`, `*.csv`) |
| MM014 | warning | неканонический lib-бутстрап — только блок `_EXTENSION_DIR`/`_LIB_DIR` (D-15) |

**Baseline** (`tools/convention_baseline.json`) — две секции одинаковой схемы `{путь: [коды]}`:

- `units` — замороженные нарушения legacy-кнопок: старые кнопки не «краснят» общий
  прогон, но и не растят долг — новые нарушения baseline не покрывает. При адаптации
  кнопки (`/mm-adopt-script`) её запись из `units` удаляется, дальше кнопка держит
  `--strict`.
- `pending_adoption` — ВРЕМЕННЫЕ допуски кнопок, физически присутствующих в рабочем
  дереве, но ещё не принятых через `/mm-adopt-script` (обычно ещё не в git). Это не
  грандфазеринг: гейт приёмки `--strict` baseline игнорирует, поэтому такая кнопка
  всё равно обязана пройти `--strict` при приёмке, а её запись из `pending_adoption`
  удаляется в том же коммите. Заводить запись в `pending_adoption` можно только для
  ещё не принятой кнопки; `--write-baseline` секцию сохраняет и её пути в `units`
  не переносит.

Тесты инструментов: `py -3 -m unittest discover -s tools/tests -q`.

## Команды MM LAB

Все команды MM LAB имеют префикс `mm-` (D-19). Канонические процедуры лежат в
`agents/commands/` (по одному файлу на команду); адаптеры агентов — в `.claude/commands/`,
`.gemini/commands/`, `.kilo/commands/` — только ссылаются на канонический файл.

| Команда | Назначение | Процедура |
| ------------------ | ----------------------------------------------------------------------- | ---------------------------------- |
| `/mm-adopt-script` | приёмка стороннего скрипта: чекер → адаптация → diff → одобрение → регистрация | `agents/commands/mm-adopt-script.md` |
| `/mm-new-button` | скаффолд новой кнопки из `templates/НоваяКнопка.pushbutton` | `agents/commands/mm-new-button.md` |
| `/mm-check` | прогон `tools/check_convention.py` по кнопке/скрипту/репозиторию | `agents/commands/mm-check.md` |
| `/mm-save-session` | сессионный коммит по шаблону + заметки Obsidian; push с подтверждением | `agents/commands/mm-save-session.md` |
| `/mm-update-repo` | безопасное обновление репозитория (fetch/pull, проверка чистоты дерева) | `agents/commands/mm-update-repo.md` |
| `/mm-doctor` | self-check: версии Revit vs поддерживаемые, vendored lib, полнота кнопок | `agents/commands/mm-doctor.md` |
| `/mm-new-compat` | добавить ветку новой версии Revit в `revit_compat.py` | `agents/commands/mm-new-compat.md` |

Фраза пользователя «сохрани сессию» = выполнить процедуру `/mm-save-session`.
Полные тексты процедур здесь не дублируются — читай файл команды в `agents/commands/`.

## Git-регламент

- Стейджинг — ТОЛЬКО пофайловый, с явными путями: `git add "путь/к/файлу"`.
  `git add .` и `git add -A` запрещены (D-17).
- Каждая сессия — отдельный коммит; в коммит попадают только файлы, затронутые
  в текущей сессии (созданные/изменённые). Сообщения коммитов — на русском.
- Push — только после подтверждения человеком (D-18); авто-push запрещён.
- Кириллические пути: статус смотри через `git -c core.quotepath=false status --porcelain`,
  иначе git экранирует русские имена файлов.

Шаблон сессионного коммита (обязательный, используется `/mm-save-session`):

```text
сессия: <краткое описание работы за сессию в одном предложении>

## Сессия
- Агент/модель: <напр. Claude Opus 4.8 / Gemini / Kilo Code>
- Дата: <YYYY-MM-DD>
- Изменено файлов: <N>

## Изменённые файлы
- <ТОЛЬКО файлы, затронутые в этой сессии (созданы/изменены)>

## Результаты
- <ключевые изменения списком>
```

## Правило graphify

Если в корне репозитория существует `graphify-out/graph.json`:

- вопросы о кодовой базе начинай с `graphify query "<вопрос>"`;
  связи между сущностями — `graphify path "<A>" "<B>"`; разбор концепта —
  `graphify explain "<концепт>"`;
- после правок кода выполняй `graphify update .` (AST-only, без API-затрат).

Если каталога `graphify-out/` нет — правило спит, ничего запускать не нужно.

## Obsidian-хранилище

Хранилище знаний: `./MMLabs_OBSIDIAN`.

- **При старте сессии:** прочитай `00-home/index.md` и `текущие приоритеты.md`;
  если задача касается конкретного модуля — соответствующую заметку из `knowledge/`.
- **При завершении сессии** (фраза «сохрани сессию») — процедура `/mm-save-session`:
  заметка в `sessions/` с датой; обновление `текущие приоритеты.md`; при принятом
  решении — заметка в `knowledge/decisions/`, при разобранном баге — в
  `knowledge/debugging/`; обновление `index.md`, если появились новые заметки;
  затем сессионный коммит по шаблону из §Git-регламент.
