# Phase 3: Конвенция скриптов MM LAB + команда приёмки/адаптации - Context

**Gathered:** 2026-07-24
**Status:** Ready for planning

<domain>
## Phase Boundary

Фаза поставляет **три связанных артефакта** + Git-регламент:

1. **Стандарт написания кода** (конвенция) для pyRevit-скриптов MM LAB — единый документ правил.
2. **Шаблон кода** (`templates/`) для старта разработки новой кнопки.
3. **Каталог `mm`-команд** (Claude/агентские slash-команды) для приёмки сторонних скриптов, скаффолда, проверки конвенции, работы с Git.

**Сквозные ограничения (зафиксированы пользователем, НЕ пересматривать):**
- Только **CPython** (`#! python3`), никакого IronPython для нового кода.
- **Мультиверсионная** поддержка Revit API обязательна.
- Конвенция работает для **разных ИИ-агентов** (Claude Code, Gemini, Kilo Code).
- Аудитория — **новички** без опыта Python/Revit API.
- Поиск параметров через `BuiltInParameter`/GUID, не `lookupParameter`; общие функции — в `lib`; обязательная шапка; README у каждого скрипта; без сторонних импортов (кроме vendored `lib/`).

**Вне scope (в другие фазы):** реальная синхронизация GSD с «Картой релизов» (этап 2 — см. Deferred); рефакторинг silent-except в существующих скриптах (отдельная задача качества).
</domain>

<decisions>
## Implementation Decisions

### 1. Мультиверсия Revit API
- **D-01:** Механизм — **единый `compat`-модуль в `lib`** (напр. `revit_compat.py`), который оборачивает версионно-зависимые вызовы в стабильные функции-хелперы. Скрипты зовут хелперы, а НЕ сырой версионный API. (Так устроено ядро pyRevit: `HOST_APP.version` + `pyrevit.compat`.) Разбрасывать `if version >=` по каждому скрипту — запрещено (противоречит правилу «общее — в lib»).
- **D-02:** Поддерживаемые версии Revit — **2020, 2022, 2024** (совпадает с docstring «Мокрых зон»). Это тест-матрица и целевые ветки `compat`.
- **D-03:** На **неподдерживаемой/неизвестной** версии Revit — **fail-fast** с понятным сообщением (перечислить поддержанные версии), а не падение с невнятной ошибкой API.
- **D-04:** `compat`-модуль обязан закрыть известные ломающие изменения (см. `<code_context>`): Units API 2021 (`ForgeTypeId`/`SpecTypeId`), `Floor.Create` 2023, `ElementId.IntegerValue→.Value` Int64 2024, pythonnet overload-обходы (`get_Parameter.__overloads__[BuiltInParameter]`, `Enum.ToObject`, `.NET` collections).

### 2. Формат стандарта для ИИ-агентов
- **D-05:** **Один канонический `AGENTS.md`** в корне репозитория — источник правды. Тонкие per-agent указатели (`CLAUDE.md`, `GEMINI.md`, `.kilocode/rules/`) коротко ссылаются на него («читай `AGENTS.md`»), НЕ дублируют. (Symlink на Windows не использовать — ссылка/синк.)
- **D-06:** **Свой Python-чекер конвенции** (напр. `tools/check_convention.py`) — машинная проверка house-правил: шапка `#! python3` + coding, отсутствие сторонних импортов, наличие README, канонический паттерн импорта из `lib`, `BuiltInParameter` вместо `lookupParameter` где применимо. Чекер запускается новичком/агентом локально и **переиспользуется командой приёмки** как гейт.
- **D-07:** Язык стандарта — **русский**; код, имена Revit API и технические термины — **английский**.

### 3. Команда приёмки/адаптации стороннего скрипта
- **D-08:** Уровень автономии — **адаптация с ревью-гейтом**: проверка чекером → правка под конвенцию → показать diff → **дождаться одобрения человека** → регистрация. Никаких тихих изменений (переносится из Phase 1).
- **D-09:** По приёмке команда заводит GSD **quick task** (лёгкий, с atomic-коммитом и записью в STATE) — не полную фазу.
- **D-10:** Целевая панель (АРХИТЕКТУРА/ИОС/КООРДИНАЦИЯ) — команда **спрашивает у пользователя с авто-подсказкой** (предлагает по содержимому скрипта), финально решает человек.
- **D-11:** Команда регистрирует принятый скрипт как полноценную `*.pushbutton` (script.py + bundle.yaml + README) и добавляет её в соответствующий `bundle.yaml`, чтобы кнопка появилась на панели.

### 4. Шаблон инициализации кода
- **D-12:** Форма — **готовая папка-скелет pushbutton** (`script.py` + `bundle.yaml` + `README.md` + место под `icon.png`). Копируешь → переименовываешь → заполняешь.
- **D-13:** Наполнение — **минимальная РАБОЧАЯ кнопка-пример** (шапка → канонический lib-бутстрап → вызов `compat` → `Transaction` → обработка ошибок → `TaskDialog`) с явными `# TODO`-метками в местах правки.
- **D-14:** Расположение — папка **`templates/` вне `MM Lab.tab`** (в корне репозитория), чтобы pyRevit не грузил шаблон как реальную кнопку. Оттуда копируют люди и `/mm-new-button`.
- **D-15:** Шаблон фиксирует **один канонический** `sys.path`-бутстрап `lib` (сейчас в репо два разных паттерна — унифицировать).

### 5. Git/GitHub-регламент и каталог mm-команд
- **D-16:** Команда завершения сессии — **`/mm-save-session`** (аналог прежнего «Сохрани сессию»). Обязательный шаблон коммита утверждён (см. `<specifics>`).
- **D-17:** В коммит попадают **только файлы, затронутые в текущей сессии** (созданные/изменённые). Стейджить пофайлово — **никаких `git add .`/`-A`**. Каждая сессия — отдельный коммит. Сообщения на русском.
- **D-18:** Политика push у `/mm-save-session` — **коммит локально, затем push с подтверждением** (не авто-push).
- **D-19:** Все агентские команды MM LAB имеют **префикс `mm-`**.
- **D-20:** Каталог команд ЭТОЙ фазы: `/mm-adopt-script`, `/mm-new-button`, `/mm-check`, `/mm-save-session`, `/mm-update-repo`, `/mm-doctor`, `/mm-new-compat`. (Точные имена/слаги подтвердить в планировании — the agent's discretion.)
- **D-21:** Из текущих `CLAUDE.md` в новый стандарт включить: правило **graphify** (вопросы о коде → сначала `graphify query`), поток **Obsidian «сохрани сессию»** (сворачивается в `/mm-save-session` + захват знаний). `userEmail`/`currentDate` — служебные, в стандарт не включать.

### Claude's Discretion
- Точные имена файлов/слагов команд и модулей (`revit_compat.py`, `check_convention.py`, слаги `mm-*`).
- Внутренняя структура `compat`-модуля и набор первых хелперов (units/ElementId/Floor/overloads).
- Формат вывода чекера и порядок правил.
- Минимальный сценарий кнопки-примера в шаблоне.
</decisions>

<canonical_refs>
## Canonical References

**Downstream agents MUST read these before planning or implementing.**

### Планирование / цели фазы
- `.planning/ROADMAP.md` §"Phase 3" — цель, объём работ (конвенция + команда приёмки), целевой релиз v260724.
- `.planning/phases/01-helper/01-CONTEXT.md` — прежние решения: shared-хелперы в `lib`, тонкие кнопки, без silent-except.
- `.planning/STATE.md` — накопленный контекст и текущая позиция (Phase 3 of 3).

### Существующий код (образцы и точки интеграции)
- `MM LAB.extension/lib/ios_common_helpers.py` — существующий shared-модуль (Phase 1/2) — паттерн для `compat`.
- `MM LAB.extension/lib/revit_ui_helpers.py` — shared UI-хелперы (`alert`).
- `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Мокрые зоны.pushbutton/script.py` — образцовая шапка + `_get_param` с `get_Parameter.__overloads__[BuiltInParameter]` (семя `compat`).
- `MM LAB.extension/MM Lab.tab/ИОС.panel/Сброс потерь.pushbutton/script.py` — альтернативный (иной) lib-бутстрап — унифицировать в шаблоне.
- `MM LAB.extension/MM Lab.tab/bundle.yaml` — регистрация панелей/кнопок (цель для команды приёмки).

### Карты кодовой базы
- `.planning/codebase/CONVENTIONS.md` — текущие naming/import/error-паттерны (основа стандарта).
- `.planning/codebase/CONCERNS.md` — тех-долг: silent-except, дублирование, дрейф vendored-зависимостей.
- `.planning/codebase/STACK.md` — pyRevit, CPython3 vs IronPython, vendored `lib/openpyxl`.
- `.planning/codebase/STRUCTURE.md` — дерево `*.tab/*.panel/*.pushbutton`, где заводить новый код.

### Инструкции агентов (свернуть в AGENTS.md)
- `CLAUDE.md` (корневой проектный) — graphify, GSD Release Map, Obsidian «сохрани сессию» + шаблон коммита.
- `.claude/CLAUDE.md` — graphify trigger, userEmail, currentDate.

### Этап 2 (deferred, читать при планировании этапа 2)
- `RELEASE_MAP/gsd_release_sync.py` — синхронизация «Карты релизов» (в разработке; интеграция — `/mm-sync-release-map`, следующая фаза).

### Внешние (deep research — Perplexity, 2026-07-24)
- Мультиверсия Revit API (pyRevit CPython, версии/ломающие изменения/pythonnet): https://www.perplexity.ai/search/178747fe-21fe-41ed-9ab0-3a6015cfff8f
- Agent-agnostic стандарт (AGENTS.md, per-agent файлы, enforcement): https://www.perplexity.ai/search/178b8fe8-c7f6-48c2-be93-c02695f9f933
- `AGENTS.md` спецификация: https://agents.md/
</canonical_refs>

<code_context>
## Existing Code Insights

### Reusable Assets
- **`lib/ios_common_helpers.py`, `lib/revit_ui_helpers.py`** — уже существующий shared-слой; `compat`-модуль встаёт рядом тем же механизмом импорта.
- **Шапка «Мокрых зон»** (`#! python3` + coding + docstring `Совместимость: Revit 2020/2022/2024` + `Зависимости: нет`) — готовый образец обязательной шапки для шаблона/стандарта.
- **`_get_param` с `get_Parameter.__overloads__[BuiltInParameter]`** — рабочий обход pythonnet-overload; первый кандидат на перенос в `compat`.
- **pyRevit `HOST_APP.version` / `pyrevit.compat`** (`IRONPY/PY3/NETCORE`, `get_elementid_value_func`) — базис для детекции версии/движка; не изобретать заново.

### Established Patterns
- **Тонкие кнопки:** `script.py` = orchestration + UI; логика — в `lib` (Phase 1).
- **Транзакции:** `Transaction.Start()`+`Commit()` в `try`, `RollBack()` в `except`; верхнеуровневый `show_error`.
- **UI-репортинг:** `TaskDialog` первично; `pyrevit.script.get_output()` вторично.
- **Порядок импортов:** stdlib → `clr.AddReference` → `Autodesk.Revit.*` → pyRevit/WinForms.

### Ключевые ломающие изменения Revit API (закрыть в compat)
- **Units 2021+:** `DisplayUnitType`/`UnitType` → `ForgeTypeId` + `UnitTypeId`/`SpecTypeId`; `UnitUtils.ConvertFromInternalUnits` сменил сигнатуру.
- **Floor 2023:** `NewFloor`/`NewSlab` удалены → `Floor.Create(doc, IList<CurveLoop>, floorTypeId, levelId, ...)`.
- **ElementId 2024:** `IntegerValue` (Int32) deprecated → `.Value` (Int64); конструктор `ElementId(long)`.
- **pythonnet vs IronPython:** int→enum каст падает на pythonnet (`Enum.ToObject(BuiltInParameter, i)`); overload через `.__overloads__`; `TaskDialog.Show` иногда требует reflection `InvokeMember`. .NET Core приходит в Revit 2025.

### Integration Points
- **Регистрация в `bundle.yaml`** (tab/panel) — команда приёмки дописывает сюда.
- **`sys.path.insert(0, LIB_DIR)`** бутстрап — унифицированная версия в шаблоне.
- **per-agent файлы** (`CLAUDE.md`/`GEMINI.md`/`.kilocode/rules/`) ссылаются на `AGENTS.md`.
</code_context>

<specifics>
## Specific Ideas

### Утверждённый шаблон коммита `/mm-save-session` (обязательный)
```
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

### Каталог `mm`-команд (эта фаза)
- `/mm-adopt-script` — приёмка/адаптация стороннего скрипта (ревью-гейт → регистрация → quick task).
- `/mm-new-button` — скаффолд кнопки из `templates/`.
- `/mm-check` — прогон `check_convention.py`.
- `/mm-save-session` — сессионный коммит по шаблону, push с подтверждением.
- `/mm-update-repo` — безопасное обновление репо (fetch/pull, проверка чистоты дерева).
- `/mm-doctor` — self-check: версия Revit vs поддерживаемые + целостность vendored `lib` + обязательные файлы кнопок.
- `/mm-new-compat` — добавить ветку новой версии Revit в `revit_compat.py`.

### Прочее
- Поддерживаемые версии Revit: **2020 / 2022 / 2024**; на прочих — fail-fast.
- «Без сторонних импортов» = **исключение для vendored `lib/`** (openpyxl/et_xmlfile) — сформулировать явно в стандарте.
</specifics>

<deferred>
## Deferred Ideas

- **Этап 2 — `/mm-sync-release-map`:** интеграция GSD с «Картой релизов» (`RELEASE_MAP/gsd_release_sync.py`, команда `Синхронизируй gsd`). Сейчас в разработке и не используется. Внедряется ВТОРЫМ этапом после основного стандарта и `mm`-команд. **Заложить как задачу-плейсхолдер (следующая фаза).**
- **Расширение `compat`** на версии 2021/2023/2025/2026 (и .NET Core 2025+) — когда парк фирмы изменится; через `/mm-new-compat`.
- **Ruff/flake8** как дополнение к своему чекеру (базовый линт стиля) — опционально позже.
- **Рефакторинг silent-except** и дедупликация в существующих скриптах под новый стандарт (задача качества, отдельно).
- Расширение shared-подхода на кнопки Архитектура/Координация (перенесено из Phase 1).

</deferred>

---

*Phase: 03-convention*
*Context gathered: 2026-07-24*
