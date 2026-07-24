# Phase 3: Конвенция правил скриптов MM LAB + команда приёмки/адаптации — Research

**Researched:** 2026-07-24
**Domain:** pyRevit-конвенция (CPython, мультиверсия Revit API), agent-agnostic стандарт (AGENTS.md), Python AST-чекер, каталог slash-команд для Claude Code / Gemini CLI / Kilo Code
**Confidence:** HIGH (ключевые механизмы проверены по официальным докам и живому коду репозитория)

<user_constraints>
## User Constraints (from CONTEXT.md)

### Locked Decisions

**Сквозные ограничения (НЕ пересматривать):**
- Только **CPython** (`#! python3`), никакого IronPython для нового кода.
- **Мультиверсионная** поддержка Revit API обязательна.
- Конвенция работает для **разных ИИ-агентов** (Claude Code, Gemini, Kilo Code).
- Аудитория — **новички** без опыта Python/Revit API.
- Поиск параметров через `BuiltInParameter`/GUID, не `lookupParameter`; общие функции — в `lib`; обязательная шапка; README у каждого скрипта; без сторонних импортов (кроме vendored `lib/`).

**1. Мультиверсия Revit API**
- **D-01:** Механизм — единый `compat`-модуль в `lib` (напр. `revit_compat.py`), оборачивающий версионно-зависимые вызовы в стабильные хелперы. Скрипты зовут хелперы, НЕ сырой версионный API. Разбрасывать `if version >=` по скриптам — запрещено.
- **D-02:** Поддерживаемые версии Revit — **2020, 2022, 2024** (тест-матрица и целевые ветки `compat`).
- **D-03:** На неподдерживаемой/неизвестной версии — **fail-fast** с понятным сообщением (перечислить поддержанные версии).
- **D-04:** `compat` обязан закрыть известные ломающие изменения: Units API 2021 (`ForgeTypeId`/`SpecTypeId`), `Floor.Create` 2023, `ElementId.IntegerValue→.Value` Int64 2024, pythonnet overload-обходы (`get_Parameter.__overloads__[BuiltInParameter]`, `Enum.ToObject`, .NET collections).

**2. Формат стандарта для ИИ-агентов**
- **D-05:** Один канонический **`AGENTS.md`** в корне — источник правды. Тонкие per-agent указатели (`CLAUDE.md`, `GEMINI.md`, `.kilocode/rules/`) коротко ссылаются на него, НЕ дублируют. Symlink на Windows не использовать.
- **D-06:** Свой Python-чекер конвенции (напр. `tools/check_convention.py`) — машинная проверка house-правил: шапка, отсутствие сторонних импортов, наличие README, канонический паттерн импорта из `lib`, `BuiltInParameter` вместо `lookupParameter` где применимо. Запускается новичком/агентом локально и переиспользуется командой приёмки как гейт.
- **D-07:** Язык стандарта — русский; код, имена Revit API и технические термины — английский.

**3. Команда приёмки/адаптации стороннего скрипта**
- **D-08:** Автономия — адаптация с ревью-гейтом: чекер → правка → diff → одобрение человека → регистрация. Никаких тихих изменений.
- **D-09:** По приёмке — GSD **quick task** (лёгкий, atomic-коммит, запись в STATE), не полная фаза.
- **D-10:** Целевую панель команда **спрашивает у пользователя с авто-подсказкой**; финально решает человек.
- **D-11:** Регистрация как полноценная `*.pushbutton` (script.py + bundle.yaml + README) + добавление в соответствующий `bundle.yaml`.

**4. Шаблон инициализации кода**
- **D-12:** Форма — готовая папка-скелет pushbutton (`script.py` + `bundle.yaml` + `README.md` + место под `icon.png`).
- **D-13:** Наполнение — минимальная РАБОЧАЯ кнопка-пример (шапка → канонический lib-бутстрап → вызов `compat` → `Transaction` → обработка ошибок → `TaskDialog`) с `# TODO`-метками.
- **D-14:** Расположение — `templates/` **вне** `MM Lab.tab` (в корне репозитория), чтобы pyRevit не грузил шаблон.
- **D-15:** Шаблон фиксирует **один канонический** `sys.path`-бутстрап `lib` (сейчас в репо два разных паттерна — унифицировать).

**5. Git/GitHub-регламент и каталог mm-команд**
- **D-16:** Команда завершения сессии — `/mm-save-session`; обязательный шаблон коммита утверждён (см. CONTEXT `<specifics>`).
- **D-17:** В коммит — только файлы текущей сессии; стейджить пофайлово, никаких `git add .`/`-A`; каждая сессия — отдельный коммит; сообщения на русском.
- **D-18:** Push-политика `/mm-save-session` — коммит локально, затем push с подтверждением (не авто-push).
- **D-19:** Все агентские команды MM LAB — префикс `mm-`.
- **D-20:** Каталог команд фазы: `/mm-adopt-script`, `/mm-new-button`, `/mm-check`, `/mm-save-session`, `/mm-update-repo`, `/mm-doctor`, `/mm-new-compat` (точные слаги — the agent's discretion).
- **D-21:** Из текущих `CLAUDE.md` включить в стандарт: правило graphify, поток Obsidian «сохрани сессию» (сворачивается в `/mm-save-session`). `userEmail`/`currentDate` — служебные, не включать.

### Claude's Discretion
- Точные имена файлов/слагов команд и модулей (`revit_compat.py`, `check_convention.py`, слаги `mm-*`).
- Внутренняя структура `compat`-модуля и набор первых хелперов (units/ElementId/Floor/overloads).
- Формат вывода чекера и порядок правил.
- Минимальный сценарий кнопки-примера в шаблоне.

### Deferred Ideas (OUT OF SCOPE)
- **Этап 2 — `/mm-sync-release-map`:** интеграция GSD с «Картой релизов» (`RELEASE_MAP/gsd_release_sync.py`). Внедряется ВТОРЫМ этапом. Заложить как задачу-плейсхолдер (следующая фаза).
- Расширение `compat` на 2021/2023/2025/2026 (и .NET Core 2025+) — через `/mm-new-compat`, когда изменится парк.
- Ruff/flake8 как дополнение к своему чекеру — опционально позже.
- Рефакторинг silent-except и дедупликация существующих скриптов — отдельная задача качества.
- Расширение shared-подхода на кнопки Архитектура/Координация (из Phase 1).
</user_constraints>

<phase_requirements>
## Phase Requirements

**Формальные REQ-ID для фазы не заведены** (в промпте — «TBD»). `.planning/REQUIREMENTS.md` генерируется из «Карты релизов» и содержит только задачи плагинов (T-101..T-301); ни один ID не относится к конвенции/командам. Требования фазы = пункты объёма работ из ROADMAP §Phase 3:

| Условный ID | Описание (из ROADMAP) | Research Support |
|----|-------------|------------------|
| CONV-STD | Документ-конвенция: BuiltInParameter/GUID вместо lookupParameter; общее — в lib; шапка `#! python3` + coding; без сторонних импортов; README у скрипта | §Architecture Patterns (AGENTS.md), §Baseline Audit, §Code Examples |
| CONV-CHECK | Проверка стороннего скрипта на соответствие конвенции | §check_convention.py, §Validation Architecture |
| CONV-ADAPT | Адаптация скрипта под конвенцию (ревью-гейт) | §mm-команды, §Common Pitfalls |
| CONV-REG | Регистрация в `bundle.yaml`, чтобы кнопка появилась на панели | §bundle.yaml layout rules |
| CONV-GSD | Создание задачи GSD при адаптации (quick task) | §GSD quick task mechanics |

Планировщику: при генерации PLAN.md использовать эти условные метки либо завести REQ-ID в ходе `/gsd-plan-phase`.
</phase_requirements>

## Summary

Фаза — «мета-инфраструктурная»: вместо кода для Revit создаются (1) стандарт (AGENTS.md + per-agent указатели), (2) `lib/revit_compat.py`, (3) `templates/` со скелетом кнопки, (4) `tools/check_convention.py`, (5) каталог из 7 `mm-*`-команд. Всё, что нужно для планирования, проверено: механика pyRevit (auto-load `<ext>.extension/lib`, семантика `layout` в bundle.yaml, выбор CPython-движка шебангом `#! python3`), ломающие изменения Revit API 2020→2022→2024 (Units/ForgeTypeId, Floor.Create, ElementId Int64), и механизмы команд у всех трёх агентов (Claude Code: `.claude/commands/*.md` ≡ skills; Gemini CLI: `.gemini/commands/*.toml`; Kilo Code: `.kilo/commands/*.md` + нативное чтение AGENTS.md).

Аудит репозитория дал критичный входной факт: **из 17 существующих кнопок конвенции соответствуют полностью ~2** (только у 7 есть `#! python3`, причём у 4 из них шебанг испорчен UTF-8 BOM; у 8 нет README; у 4 нет bundle.yaml; свежие `IFC_Двери`/`IFC_Окна` — живой пример «стороннего скрипта», который надо принимать через `/mm-adopt-script`). Значит чекер обязан иметь **baseline/legacy-режим** (список grandfathered-кнопок), иначе он «красный» с первого дня и бесполезен как гейт приёмки.

Вторая ключевая находка: в репо **две папки `lib`** (корневая `lib/` = vendored openpyxl/et_xmlfile; `MM LAB.extension/lib/` = first-party хелперы, автоматически добавляется pyRevit в sys.path) и **три варианта бутстрапа** в скриптах, в двух из которых переменная `EXTENSION_ROOT` на самом деле указывает на корень репозитория. Канонический бутстрап шаблона (D-15) должен явно разводить эти два каталога.

**Primary recommendation:** AGENTS.md в корне = полный текст конвенции (русский); `CLAUDE.md` → одна строка `@AGENTS.md` + Claude-специфика, `GEMINI.md` → `@AGENTS.md` (memory import), Kilo читает AGENTS.md нативно. Канонические процедуры `mm-*` — по одному markdown-файлу на команду в одном каталоге, тонкие адаптеры в `.claude/commands/`, `.gemini/commands/*.toml`, `.kilo/commands/`. Чекер — stdlib-only (`ast` + `sys.stdlib_module_names`, Python ≥3.10), с `--json`, `--strict` и baseline.

## Project Constraints (from CLAUDE.md)

Директивы действующих CLAUDE.md (root + `.claude/CLAUDE.md`), обязательные для планов фазы:

1. **graphify:** вопросы о кодовой базе — сначала `graphify query`, когда существует `graphify-out/graph.json` (сейчас `graphify-out/` НЕ существует — правило спит; в AGENTS.md переносится с этим условием, D-21). После правок кода — `graphify update .`.
2. **GSD Release Map:** команда «Синхронизируй gsd» → `python3 RELEASE_MAP/gsd_release_sync.py check|sync-docs`; статусы задач: Не начато → В работе → Готово → Релиз. (Интеграция в `mm-*` — deferred, этап 2.)
3. **Obsidian vault `./MMLabs_OBSIDIAN`:** при «сохрани сессию» — заметка в `sessions/`, обновление `текущие приоритеты.md`, `knowledge/decisions|debugging`, `index.md`; коммит **только файлов текущей сессии** по утверждённому шаблону. Этот поток сворачивается в `/mm-save-session` (D-16, D-21).
4. `userEmail`/`currentDate` — служебный контекст, в стандарт не включать (D-21).

Планы не должны противоречить этим директивам; наоборот, AGENTS.md их поглощает, а root CLAUDE.md становится тонким указателем (см. Open Questions Q1 о судьбе текущего содержимого CLAUDE.md).

## Standard Stack

Фаза не тянет новых runtime-зависимостей. «Стек» = форматы и встроенные инструменты:

### Core
| Инструмент/формат | Версия | Назначение | Почему стандарт |
|---------|---------|---------|--------------|
| pyRevit extension model (`*.tab/*.panel/*.pushbutton`, `bundle.yaml`) | у пользователей предположительно pyRevit 5.x, pythonnet 3.x [ASSUMED — см. A1] | Регистрация кнопок, layout | Уже основа репо; layout-директива управляет видимостью/порядком [CITED: pyrevit1.readthedocs.io/en/latest/creatingexts.html] |
| CPython через `#! python3` (первая строка) | pyRevit CPython engine | Движок новых скриптов | Официальный механизм выбора движка [VERIFIED: learnrevitapi.com, pyRevit docs] |
| Python (локальный, для tools/) | 3.14.2 / 3.13.7 установлены [VERIFIED: `python --version` на машине] | `check_convention.py`, unittest | `sys.stdlib_module_names` (3.10+) для белого списка stdlib |
| `ast` (stdlib) | stdlib | Разбор script.py без исполнения | Безопасно для стороннего кода (никаких import/exec) |
| `unittest` (stdlib) | stdlib | Тесты чекера | Ноль зависимостей — в духе «без сторонних импортов» |
| AGENTS.md (agents.md spec) | открытый формат, plain Markdown | Канонический стандарт | 20+ инструментов; discovery в корне репо [CITED: agents.md] |
| Claude Code commands/skills | текущая версия CC | `/mm-*` для Claude | «Custom commands have been merged into skills» — `.claude/commands/deploy.md` и `.claude/skills/deploy/SKILL.md` равнозначно создают `/deploy` [CITED: code.claude.com/docs/en/slash-commands] |
| Gemini CLI custom commands | текущая версия | `/mm-*` для Gemini | `.gemini/commands/*.toml`, ключи `prompt`(req)/`description`, плейсхолдер `{{args}}`, шелл `!{...}` [CITED: geminicli.com/docs/cli/custom-commands] |
| Kilo Code commands + AGENTS.md | текущая версия | `/mm-*` для Kilo | `.kilo/commands/*.md` (авто-миграция из legacy `.kilocode/workflows/`); AGENTS.md в корне читается нативно и грузится ПЕРВЫМ, до `.kilocode/rules/` [CITED: kilo.ai/docs/customize/workflows; GitHub Kilo-Org kilo-docs agents-md] |

### Supporting
| Инструмент | Назначение | Когда |
|---------|---------|-------------|
| `pyrevit.compat` (`get_elementid_value_func`, `NETCORE`, `_get_revit_version`) | Образец для `revit_compat.py` — не изобретать детекцию | Ссылаться как на паттерн; наш compat — самодостаточный [VERIFIED: raw.githubusercontent.com pyrevitlabs/pyRevit compat.py] |
| git 2.55 [VERIFIED: локально] | `/mm-save-session`, `/mm-update-repo` | `core.quotepath=false` для кириллических путей |
| Node 20.20 [VERIFIED: локально] | GSD tooling (quick task) | Только для GSD-интеграции |

### Alternatives Considered
| Вместо | Можно | Tradeoff |
|------------|-----------|----------|
| Свой AST-чекер | ruff/flake8 + свои плагины | House-правила (шапка, bundle.yaml, README, layout) линтеры не покрывают; ruff — deferred по CONTEXT |
| unittest | pytest | pytest удобнее, но = pip-зависимость; для новичка `python -m unittest` без установки надёжнее |
| Тонкие per-agent адаптеры команд | Полные копии процедур в 3 форматах | Копии дрейфуют; адаптер «прочитай канонический файл и выполни» — единый источник правды (в духе D-05) |
| `.claude/commands/mm-*.md` | `.claude/skills/mm-*/SKILL.md` | Равнозначно по функции; `commands/` изолирует mm-* от GSD-managed `skills/` (там gsd-file-manifest.json и 60+ gsd-скиллов) |

**Installation:** ничего устанавливать не нужно (всё stdlib/уже в репо). Установленные Python 3.13/3.14, git, Node проверены на машине.

## Architecture Patterns

### Recommended Project Structure (deliverables фазы)

```
MMLab_TOOLS/
├── AGENTS.md                          # КАНОНИЧЕСКИЙ стандарт (RU, полный текст конвенции)
├── CLAUDE.md                          # тонкий: @AGENTS.md + Claude-специфика (GSD, skills)
├── GEMINI.md                          # тонкий: @AGENTS.md (Gemini memory import)
├── agents/
│   └── commands/                      # канонические процедуры mm-* (по 1 md на команду)
│       ├── mm-adopt-script.md
│       ├── mm-new-button.md
│       ├── mm-check.md
│       ├── mm-save-session.md
│       ├── mm-update-repo.md
│       ├── mm-doctor.md
│       └── mm-new-compat.md
├── .claude/commands/mm-*.md           # адаптеры Claude (frontmatter + «выполни agents/commands/…»)
├── .gemini/commands/mm-*.toml         # адаптеры Gemini (prompt → канонический файл, {{args}})
├── .kilo/commands/mm-*.md             # адаптеры Kilo (то же)
├── templates/
│   └── НоваяКнопка.pushbutton/        # D-12..D-15; ВНЕ MM LAB.extension → pyRevit не грузит
│       ├── script.py                  # рабочий пример: шапка → бутстрап → compat → Transaction → TaskDialog
│       ├── bundle.yaml                # title/tooltip (ru), author + TODO
│       ├── README.md                  # каркас по образцу «Мокрых зон»
│       └── icon.png                   # плейсхолдер (или README-инструкция «положи icon.png»)
├── tools/
│   ├── check_convention.py            # stdlib-only чекер
│   ├── convention_baseline.json       # grandfathered-кнопки (legacy-исключения)
│   └── tests/                         # unittest + fixtures (good_button/, bad_button/)
└── MM LAB.extension/lib/
    └── revit_compat.py                # D-01..D-04
```

Имена файлов/папок — the agent's discretion (зафиксированы решениями только префикс `mm-`, `templates/` в корне, compat в `MM LAB.extension/lib`).

### Pattern 1: Двухуровневая схема «AGENTS.md + тонкие указатели» (D-05)

**Факты по агентам (проверено):**
- **Claude Code НЕ читает AGENTS.md** (подтверждено доками Anthropic, май 2026). Официальный мост — импорт в CLAUDE.md: строка `@AGENTS.md` разворачивается при старте сессии, как будто текст написан в CLAUDE.md. Symlink на Windows требует прав — не использовать (совпадает с D-05). [VERIFIED: web, несколько источников + docs]
- **Gemini CLI:** контекст-файл по умолчанию GEMINI.md; поддерживает импорты `@file.md` (Memory Import Processor) И настройку `context.fileName`/`contextFileName` в `.gemini/settings.json` (можно `["AGENTS.md","GEMINI.md"]`). Свежие версии заявлены на agents.md как поддерживающие AGENTS.md нативно. Надёжный минимум без предположений о версии: `GEMINI.md` с `@AGENTS.md`. [CITED: github.com/google-gemini/gemini-cli docs/cli/gemini-md.md]
- **Kilo Code:** AGENTS.md в корне workspace читается нативно и загружается ПЕРВЫМ (до `.kilocode/rules/*.md`); `.kilocode/rules/` остаётся поддержанным (backward compatible). То есть для Kilo отдельный указатель не обязателен; можно добавить `.kilocode/rules/00-mmlab.md` в одну строку для старых версий. [CITED: kilo docs / DeepWiki Kilo-Org]

**Следствие для планирования:** AGENTS.md может быть единственным местом текста конвенции; три «указателя» суммарно ~10 строк. Claude-специфику (GSD-команды, ссылки на skills) оставить в CLAUDE.md ПОД строкой импорта.

### Pattern 2: Канонические процедуры mm-команд + per-agent адаптеры (D-19, D-20)

Механика по агентам:

| Агент | Файл | Формат | Вызов |
|-------|------|--------|-------|
| Claude Code | `.claude/commands/mm-check.md` | Markdown + YAML frontmatter (`description`, `argument-hint`, `allowed-tools`, `model`); `$ARGUMENTS`/`$1..$n`; `!`-bash; `@file` | `/mm-check <args>` |
| Gemini CLI | `.gemini/commands/mm-check.toml` | TOML: `prompt` (обязателен), `description`; `{{args}}`; `!{cmd}` шелл-вставки | `/mm-check <args>` (плоское имя файла с дефисами; подпапки дают `:`-неймспейс — НЕ использовать) |
| Kilo Code | `.kilo/commands/mm-check.md` | Markdown + optional frontmatter (`description`, `agent`, `model`) | `/mm-check` (имя файла без `.md`); legacy `.kilocode/workflows/` авто-мигрирует |

Анти-дрейф: тело каждого адаптера — 2–4 строки: «Прочитай `agents/commands/mm-check.md` и выполни процедуру. Аргументы: …». Канонический файл содержит полную процедуру (RU), включая гейты D-08/D-10/D-17/D-18.

Примечание: в Claude Code custom commands объединены со skills — оба пути валидны; `commands/` выбран, чтобы не смешиваться с GSD-managed `.claude/skills/` (у GSD собственный манифест `.claude/gsd-file-manifest.json`, обновляемый `/gsd-update`).

### Pattern 3: Структура `revit_compat.py` (D-01..D-04, рекомендация — детали discretion)

```python
#! python3
# -*- coding: utf-8 -*-
"""Совместимость Revit API 2020/2022/2024 для скриптов MM LAB. Зависимости: нет."""
SUPPORTED_VERSIONS = (2020, 2022, 2024)

def get_revit_version():      # int, напр. 2024
    # каскад: builtins.__revit__ → __revit__ в globals вызвавшего скрипта (передан явно)
    # → pyrevit.HOST_APP (допустимый импорт хост-платформы)
    ...

def require_supported_version(command_name):   # D-03 fail-fast
    # если версии нет в SUPPORTED_VERSIONS → TaskDialog со списком поддержанных + raise SystemExit
    ...

# --- параметры (закрывает pythonnet-обходы, D-04) ---
def get_parameter(element, built_in_parameter, *fallback_names): ...
    # каскад: element.get_Parameter(bip) → get_Parameter.__overloads__[BuiltInParameter](bip)
    # → BIP→ParameterElement.Definition.Name→LookupParameter (кеш) → fallback_names
def get_shared_parameter(element, guid): ...       # get_Parameter(Guid) для общих параметров

# --- ElementId 2024 Int64 ---
def element_id_value(element_id): ...              # .Value → .IntegerValue fallback
def make_element_id(int_value): ...                # ElementId(Int64) на 2024, ElementId(int) ранее

# --- Units: 2020 = DisplayUnitType, 2022/2024 = ForgeTypeId/UnitTypeId ---
def convert_from_internal(value, unit_key): ...    # unit_key: "mm" | "m" | "m2" | ...
def convert_to_internal(value, unit_key): ...

# --- Floor: 2020 = doc.Create.NewFloor(CurveArray,..), 2022/2024 = Floor.Create(doc, IList[CurveLoop],..) ---
def create_floor(doc, curve_loops, floor_type_id, level_id): ...

# --- pythonnet interop ---
def to_net_list(items, net_type): ...              # System.Collections.Generic.List[T]
def enum_from_int(enum_type, int_value): ...       # Enum.ToObject
def iter_count(net_or_py_seq): ...                 # len()-безопасный подсчёт (кейс .Count на IList, коммит 3fcf888)
```

Семена уже в репо: `ios_common_helpers._bip_to_lookup_name`+`get_parameter` (BIP→имя→LookupParameter, кеш) и `Мокрые зоны._get_param` (`__overloads__`-обход). **В репо два конкурирующих обхода get_Parameter — compat обязан выбрать один канонический каскад**, а `ios_common_helpers` со временем делегировать в compat (миграция — за пределами обязательного объёма, но стоит задача «compat не создаёт третий вариант»).

Проверенные факты для веток:
- `ElementId.IntegerValue` (Int32) deprecated в 2024; `.Value` (Int64) — замена; конструктор `ElementId(System::Int64)` [VERIFIED: Autodesk blog «What's New in the Revit 2024 API», forums]
- Units: `DisplayUnitType` deprecated в 2021, заменён `ForgeTypeId`+`UnitTypeId`; `UnitUtils.ConvertFromInternalUnits(double, ForgeTypeId)` с 2021; в 2022 старый DUT-API удалён [VERIFIED: revitapidocs 2022, archi-lab.net «handling the revit 2022 unit changes»] → для матрицы 2020/2022/2024 ветка ровно одна: 2020 = DUT, 2022+ = ForgeTypeId.
- Floors: `Document.NewFloor()/NewSlab()` deprecated в 2022, удалены в 2023; `Floor.Create(Document, IList<CurveLoop>, ElementId, ElementId)` с 2022; профиль сменился CurveArray→CurveLoop [VERIFIED: revitapidocs 2022/2024, Autodesk forums, pyRevit forum]
- `pyrevit.compat` даёт готовые образцы: `get_elementid_value_func()` («Value» для пост-2023, иначе «IntegerValue»), `NETCORE` (Revit 2025+ / .NET 8 — вне матрицы, deferred), `_get_revit_version()` каскад UIApplication→Application→ControlledApplication [VERIFIED: исходник compat.py master]

### Pattern 4: Канонический lib-бутстрап шаблона (D-15)

Сейчас в репо ТРИ варианта, и в двух `EXTENSION_ROOT` фактически указывает на корень репо (4×`..` от pushbutton), а не на `MM LAB.extension`:
- «Мокрые зоны»: 4×`..` → корень репо → `lib/` (vendored!), `insert(0)`; первопартийный импорт `revit_ui_helpers` при этом работает только потому, что pyRevit сам добавил `MM LAB.extension/lib` в sys.path.
- «Сброс потерь»: 3×`dirname` → `MM LAB.extension` → `lib/` (first-party), `append`.
- «Экспорт ПСО»: 4×`..` → корень репо → `lib/` (vendored, для openpyxl), `insert(0)`.

**Факт pyRevit:** папка `lib` в корне UI-extension — встроенная фича, добавляется в sys.path всем командам расширения при старте; изменения в `lib` требуют Reload/перезапуска Revit (парсится на старте) [VERIFIED: discourse.pyrevitlabs.io/t/7764 + pyrevit1.readthedocs]. `.lib`-extensions добавляются в sys.path всех расширений.

**Рекомендация для шаблона (одна каноническая форма, имена честные):**
```python
import os, sys
_SCRIPT_DIR = os.path.dirname(__file__)
# pushbutton → panel → tab → MM LAB.extension
_EXTENSION_DIR = os.path.normpath(os.path.join(_SCRIPT_DIR, "..", "..", ".."))
_LIB_DIR = os.path.join(_EXTENSION_DIR, "lib")            # first-party (revit_compat и др.)
if os.path.isdir(_LIB_DIR) and _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)
```
Vendored-каталог (корневой `lib/` с openpyxl) в шаблон НЕ включать; он нужен единицам скриптов — подключать через хелпер в compat (напр. `revit_compat.ensure_vendor_lib()`), вызываемый после бутстрапа. Это разводит «first-party lib» и «vendored lib» терминологически и в конвенции.

### Pattern 5: Архитектура `tools/check_convention.py` (D-06)

- **Запуск:** обычный CPython ≥3.10 (на машине 3.13/3.14), без Revit: только `ast` + файловая система. `py -3 tools/check_convention.py <paths>|--all [--strict] [--json] [--baseline tools/convention_baseline.json]`.
- **Единица проверки:** папка `*.pushbutton` (script.py + bundle.yaml + README.md + panel-layout) — не одиночный файл; плюс режим «сырой script.py» для приёмки ДО скаффолда.
- **Правила (коды, рекомендация):**

| Код | Проверка | Severity | Основание |
|-----|----------|----------|-----------|
| MM001 | строка 1 == `#! python3` | error | конвенция; выбор движка pyRevit |
| MM002 | строка 2 == `# -*- coding: utf-8 -*-` | error | конвенция (PEP 263: coding — строка 1–2) |
| MM003 | файл без UTF-8 BOM | error | 4 кнопки ИОС имеют BOM ПЕРЕД шебангом (найдено аудитом: `EF BB BF #! python3`) |
| MM004 | docstring с «Совместимость: Revit 2020 / 2022 / 2024» и «Зависимости:» | warning | образец «Мокрых зон» |
| MM005 | bundle.yaml существует, есть `title:`/`tooltip:` | error | D-11 |
| MM006 | README.md существует | error | конвенция |
| MM007 | кнопка есть в panel `bundle.yaml layout`; все layout-записи имеют папки | error | иначе кнопка не видна на pyRevit ≤5.x; в tab-layout уже есть орфан «ВОР» без панели |
| MM008 | импорты только из белого списка: stdlib (`sys.stdlib_module_names`) + `clr`,`System`,`Autodesk`,`pyrevit` + модули `MM LAB.extension/lib/*.py` + vendored top-level (`openpyxl`,`et_xmlfile`) | error | «без сторонних импортов, кроме vendored lib» |
| MM009 | нет `from X import *` | error | IFC_Двери: `from Autodesk.Revit.DB import *` |
| MM010 | `LookupParameter("строковый литерал")` → предложить `revit_compat.get_parameter`/GUID | warning | D-06 «где применимо»; в lib есть легитимный fallback |
| MM011 | голый `except:` | warning | рефакторинг silent-except deferred → не error |
| MM012 | `from pyrevit import forms` под CPython | warning | комментарий в «Экспорт ПСО»: pyrevit.forms не работает под CPython3 |
| MM013 | мусор в папке кнопки (`__pycache__/`, `*.csv`, `.vs/`) | warning | найдено в «Мокрых зонах» |

- **Baseline:** JSON-список кнопок с допущенными кодами (17 текущих кнопок «замораживаются»); новые/адаптируемые проверяются `--strict` (все правила = error, baseline игнорируется). Гейт приёмки D-08 = `--strict`.
- **Вывод:** человекочитаемый (RU) + `--json` для агентов; exit 0 = чисто, 1 = нарушения, 2 = ошибка использования. Обязательно `sys.stdout.reconfigure(encoding="utf-8")` — консоль Windows cp1251/cp866 иначе падает на кириллице.
- **bundle.yaml парсинг:** PyYAML в stdlib нет; файлы примитивны (title/tooltip/layout) → построчный парсер ограниченной схемы, а не зависимость. Ограничение задокументировать в docstring чекера.

### Pattern 6: Правила bundle.yaml / layout (для `/mm-adopt-script`, D-11)

Проверено по докам и живому репо:
- Panel `bundle.yaml` → ключ `layout:` — список имён папок БЕЗ суффикса `.pushbutton`; порядок = порядок кнопок; `---` (3+ дефиса) — разделитель; `>>>` (3+) — всё ниже уходит в slide-out [CITED: pyrevit docs/Notion Bundle Layout, readthedocs].
- **Кнопка, не указанная в layout, на pyRevit ≤5.x НЕ отображается** (исторически документированное поведение); в pyRevit 6.4.0 с новым launcher кнопки показываются независимо от layout (регресс/изменение, обсуждается) [VERIFIED: discourse.pyrevitlabs.io/t/10078]. → конвенция: регистрация в layout ОБЯЗАТЕЛЬНА (портируемо между версиями), чекер MM007.
- Tab `bundle.yaml` — тот же `layout:` для панелей; сейчас содержит орфан `ВОР` (панели нет) и хвостовые пробелы — pyRevit терпит, чекер должен ловить.
- Command `bundle.yaml`: `title:` (локализуемый `ru:`/`en_us:`, `\n` для переноса), `tooltip:`, `author:` — по образцам репо. `bundle.yaml` кнопки опционален для pyRevit (без него имя = имя папки: живой пример «Проверка зон»), но обязателен по конвенции.
- После правки bundle.yaml/lib нужен **pyRevit Reload** (структура парсится на старте) — команда приёмки должна писать об этом пользователю.

### Pattern 7: GSD quick task из `/mm-adopt-script` (D-09)

Механика в репо (по факту `.planning/quick/260709-jko-*/`): папка `.planning/quick/<id>-<slug>/` с `<id>-PLAN.md` + `<id>-SUMMARY.md`, строка в таблице STATE.md «Quick Tasks Completed». В Claude Code канонично делегировать в `/gsd-quick` (skill существует). Для Gemini/Kilo (где GSD-скиллов нет) процедура канонического файла описывает создание тех же артефактов вручную по шаблону. Планировщику: описать оба пути в `agents/commands/mm-adopt-script.md`.

### Anti-Patterns to Avoid
- **`EXTENSION_ROOT`, указывающий на корень репо** — переименовать честно (`_REPO_ROOT`/`_EXTENSION_DIR`), в шаблоне не воспроизводить.
- **Третий вариант обхода get_Parameter** — compat консолидирует два существующих, скрипты зовут только compat.
- **Дублирование текста конвенции в CLAUDE.md/GEMINI.md/.kilocode** — только указатели (D-05).
- **`git add .` / `-A` в mm-командах** — запрещено D-17; только пофайловый стейджинг.
- **Шаблон внутри `MM Lab.tab`** — pyRevit загрузит его как кнопку; только корневой `templates/` (D-14).
- **Полные процедуры в 3 форматах команд** — дрейф гарантирован; тонкие адаптеры.

## Don't Hand-Roll

| Проблема | Не строить | Использовать | Почему |
|---------|-------------|-------------|-----|
| Детекция версии/движка Revit | свой парсер окружения | паттерн `pyrevit.compat._get_revit_version` (каскад UIApplication→Application→ControlledApplication) + `builtins.__revit__` | краевые случаи (startup, ControlledApplication) уже решены в pyRevit |
| ElementId 64-bit | ручные if по всему коду | паттерн `get_elementid_value_func()` из pyrevit.compat, обёрнутый в свой compat | D-01: версионные if только внутри compat |
| Классификация stdlib-импортов в чекере | свой список из головы | `sys.stdlib_module_names` (Python 3.10+) | полнота и актуальность по версии интерпретатора |
| Парсинг Python-кода в чекере | regex по исходнику | модуль `ast` | regex ловит строки в комментариях/литералах; ast даёт точные узлы Import/ImportFrom/ExceptHandler/Call |
| Slash-команды агентов | свой диспетчер | нативные механизмы: `.claude/commands`, `.gemini/commands/*.toml`, `.kilo/commands` | поддерживаются вендорами, автодискавери, аргументы из коробки |
| Единый стандарт для агентов | свой формат правил | AGENTS.md (agents.md spec) | 20+ инструментов читают; Kilo — нативно; Gemini/Claude — импортом |
| Квик-таски приёмки | свой трекинг | GSD `/gsd-quick` артефакты (`.planning/quick/`, STATE-таблица) | уже принятый процесс репо (прецедент 260709-jko) |

**Key insight:** вся «версионная боль» Revit уже прожита в pyRevit core и в двух рабочих обходах этого репо — фаза не решает новые задачи Revit API, а консолидирует проверенные решения в один модуль + фиксирует их текстом.

## Baseline Audit (входное состояние — 17 кнопок)

Факты для планировщика (проверено скриптом по репо, 2026-07-24):

| Метрика | Значение |
|---------|----------|
| `#! python3` первой строкой | 7/17 (из них 4 — с BOM перед шебангом: Доп расход 0/1, Конфузор-Диффузор, Приточный по классификации) |
| Без шебанга (IronPython по умолчанию) | 10/17, включая свежие IFC_Двери/IFC_Окна |
| Нет README.md | 8/17 |
| Нет bundle.yaml | 4/17 (IFC_Двери, IFC_Окна, Кривые панели (Импост), Проверка зон) |
| Не зарегистрированы в panel layout | IFC_Двери, IFC_Окна, Кривые панели (Импост) — на pyRevit ≤5.x не видны |
| Орфаны в layout | tab bundle.yaml: `ВОР` без папки `ВОР.panel` |
| Мусор в папках кнопок | `__pycache__/`, `wet_zones_report.csv` («Мокрые зоны»), `.vs/` |
| Нарушения в IFC_Двери (живой кейс приёмки) | нет шебанга; `from Autodesk.Revit.DB import *`; `pyrevit.forms/script` (CPython-несовместимо); `LookupParameter("GP_23_Назначение")`; голые `except: pass`; нет bundle.yaml/README/icon.png; icon-файл `icons8-дверь-100.png` |

**Выводы:** (1) baseline-режим чекера обязателен; (2) IFC_Двери/IFC_Окна — идеальный приёмочный сценарий для UAT `/mm-adopt-script` (но их фактическая адаптация — исполнение команды, не обязательный объём фазы); (3) правило MM003 (BOM) добавить в конвенцию явно — реальный класс дефекта.

## Common Pitfalls

### Pitfall 1: UTF-8 BOM перед шебангом
**Что ломается:** `EF BB BF` перед `#! python3` (4 кнопки ИОС; редакторы Windows «Сохранить как UTF-8» добавляют BOM). На текущей среде пользователей скрипты, судя по классу ошибок из quick task 260709 (pythonnet TypeError), исполняются CPython — то есть их pyRevit BOM терпит; но это недокументированное поведение и межверсионный риск.
**Как избежать:** конвенция «UTF-8 без BOM»; чекер MM003 сверяет первые 3 байта; шаблон и адаптация перезаписывают файл без BOM.

### Pitfall 2: layout-семантика зависит от версии pyRevit
**Что ломается:** на pyRevit ≤5.x кнопка вне `layout` скрыта; на 6.4.0 (новый launcher) — показывается всё. Команда приёмки, полагающаяся на «не в layout = не видно», ошибётся на новых версиях; и наоборот.
**Как избежать:** конвенция — регистрация в layout всегда обязательна (MM007); `/mm-doctor` может сообщать версию pyRevit, если доступна.

### Pitfall 3: `__revit__` внутри lib-модулей
**Что происходит:** `__revit__` инжектится pyRevit в скоуп скрипта; `ios_common_helpers.get_document` обращается к нему из lib-модуля и в проде работает (видимо, через builtins на их движке) — но это НЕ гарантированный контракт для всех движков/версий.
**Как избежать:** compat детектит версию каскадом с `getattr(builtins, "__revit__", None)` и фолбэками; конвенция для новых скриптов — получать `doc`/`uidoc` в script.py и передавать в функции параметрами.

### Pitfall 4: pyRevit парсит расширение на старте
**Что ломается:** новые кнопки/правки `lib` не подхватываются до Reload; «команда выполнена, а кнопки нет» — ложный баг-репорт новичка.
**Как избежать:** `/mm-adopt-script` и `/mm-new-button` в финальном сообщении: «Сделай pyRevit Reload (или перезапусти Revit)». [VERIFIED: forum 7764]

### Pitfall 5: кириллица в консоли и в git
**Что ломается:** (а) `print` кириллицы из чекера на cp866/cp1251-консоли → UnicodeEncodeError; (б) `git status` экранирует кириллические пути (`"\320\220..."`) при `core.quotepath=true` (по умолчанию) — агенты неверно парсят имена файлов.
**Как избежать:** чекер — `sys.stdout.reconfigure(encoding="utf-8")`; mm-команды с git — `git -c core.quotepath=false status --porcelain` (или `-z`).

### Pitfall 6: pyrevit.forms под CPython
**Что ломается:** `pyrevit.forms` не работает под CPython3 (зафиксировано комментарием в «Экспорт ПСО»); сторонние скрипты (IFC_*) активно его используют → слепая простановка шебанга без замены forms ломает кнопку.
**Как избежать:** адаптация в `/mm-adopt-script`: `pyrevit.forms.*` → WinForms-паттерны репо (`SaveFileDialog`, свои Form-диалоги) или `TaskDialog`; чекер MM012 предупреждает.

### Pitfall 7: «без сторонних импортов» ≠ «без pyrevit»
**Что ломается:** слишком строгий белый список импортов забанит `pyrevit` (host-платформа) или `clr/System` — и весь репо «красный».
**Как избежать:** белый список явно: stdlib + `clr` + `System` + `Autodesk` + `pyrevit` + first-party lib + vendored (`openpyxl`, `et_xmlfile`); формулировку «исключение для vendored lib/» внести в стандарт (из CONTEXT `<specifics>`).

### Pitfall 8: чекер на CPython 3.13 vs pyRevit-движок
**Что ломается:** локальный `ast.parse`/`py_compile` под 3.13 примет синтаксис, которого нет в движке pyRevit (у пользователей CPython-движок из состава pyRevit — версия зависит от установки: ~3.8 у pyRevit 4.8, ~3.12 у pythonnet3-движка pyRevit 5). Обратное тоже: старый код валиден.
**Как избежать:** конвенция — консервативный синтаксис (без match/walrus в скриптах кнопок — рекомендация); зафиксировать в стандарте, что чекер = статический гейт, runtime-истина = Revit UAT.

### Pitfall 9: Gemini-неймспейсы против префикса mm-
**Что ломается:** файл `.gemini/commands/mm/check.toml` даёт команду `/mm:check` (двоеточие), а не `/mm-check` — каталог разъезжается по именам между агентами.
**Как избежать:** у Gemini — плоские файлы `mm-check.toml` (дефис в имени файла = дефис в команде).

## Code Examples

### Обязательная шапка (канон, по «Мокрым зонам»)
```python
#! python3
# -*- coding: utf-8 -*-
"""Название кнопки

Что делает, кратко.

Совместимость: Revit 2020 / 2022 / 2024
Зависимости: нет
"""
__title__ = "Две\nстроки"
__author__ = "GENPRO LAB"
```
[VERIFIED: живой образец репо]

### Fail-fast по версии (D-03) в начале main()
```python
from revit_compat import require_supported_version
require_supported_version(COMMAND_NAME)   # TaskDialog + SystemExit на 2021/2023/2025/...
```

### Ветвление Units внутри compat (единственное место с if по версии)
```python
# Source: revitapidocs 2022 ConvertFromInternalUnits(Double, ForgeTypeId); archi-lab.net (2022 unit changes)
if REVIT_VERSION <= 2020:
    from Autodesk.Revit.DB import DisplayUnitType
    _UNIT = {"mm": DisplayUnitType.DUT_MILLIMETERS, "m": DisplayUnitType.DUT_METERS}
else:  # 2022, 2024
    from Autodesk.Revit.DB import UnitTypeId
    _UNIT = {"mm": UnitTypeId.Millimeters, "m": UnitTypeId.Meters}

def convert_from_internal(value, unit_key):
    from Autodesk.Revit.DB import UnitUtils
    return UnitUtils.ConvertFromInternalUnits(value, _UNIT[unit_key])
```

### ElementId (по образцу pyrevit.compat.get_elementid_value_func)
```python
# Source: pyrevitlabs/pyRevit pyrevitlib/pyrevit/compat.py (master)
def element_id_value(element_id):
    try:
        return element_id.Value          # Revit 2024+: Int64
    except AttributeError:
        return element_id.IntegerValue   # 2020/2022
```

### AST-чекер: импорты и голые except (ядро MM008/MM009/MM011)
```python
import ast, sys
ALLOWED_ROOTS = set(sys.stdlib_module_names) | {"clr", "System", "Autodesk", "pyrevit"} \
    | FIRST_PARTY_LIB | VENDORED  # lib/*.py и openpyxl/et_xmlfile

tree = ast.parse(source, filename=str(path))   # '#! python3' — обычный комментарий, парсится
for node in ast.walk(tree):
    if isinstance(node, ast.ImportFrom):
        if node.names[0].name == "*":
            report("MM009", node.lineno)                      # wildcard
        root = (node.module or "").split(".")[0]
        if node.level == 0 and root and root not in ALLOWED_ROOTS:
            report("MM008", node.lineno, root)
    elif isinstance(node, ast.Import):
        for alias in node.names:
            root = alias.name.split(".")[0]
            if root not in ALLOWED_ROOTS:
                report("MM008", node.lineno, root)
    elif isinstance(node, ast.ExceptHandler) and node.type is None:
        report("MM011", node.lineno)                          # bare except
```

### Тонкий указатель CLAUDE.md (D-05)
```markdown
@AGENTS.md

<!-- Ниже — только специфичное для Claude Code (GSD-команды, skills). -->
```
[VERIFIED: механика @import в CLAUDE.md — доки Anthropic; «Claude Code не читает AGENTS.md напрямую»]

### Адаптер Gemini `.gemini/commands/mm-check.toml`
```toml
# Source: geminicli.com/docs/cli/custom-commands
description = "MM LAB: проверить скрипт/кнопку на соответствие конвенции"
prompt = """
Прочитай файл agents/commands/mm-check.md и выполни описанную процедуру.
Аргументы пользователя: {{args}}
"""
```

### Адаптер Kilo `.kilo/commands/mm-check.md`
```markdown
---
description: "MM LAB: проверить скрипт/кнопку на соответствие конвенции"
---
Прочитай файл agents/commands/mm-check.md и выполни описанную процедуру. Аргументы: всё после /mm-check.
```
[CITED: kilo.ai/docs/customize/workflows — `.kilo/commands/submit-pr.md` → `/submit-pr`]

### Panel bundle.yaml — регистрация кнопки (что дописывает /mm-adopt-script)
```yaml
layout:
  - Мокрые зоны
  - Экспорт ПСО
  - НоваяКнопка        # ← имя папки БЕЗ .pushbutton; не в layout ⇒ скрыта на pyRevit ≤5.x
```

## State of the Art

| Старый подход | Текущий | Когда сменилось | Значение для фазы |
|--------------|------------------|--------------|--------|
| `DisplayUnitType`/`UnitUtils(…, DUT)` | `ForgeTypeId` + `UnitTypeId`/`SpecTypeId` | deprecated 2021, удалено 2022 | ветка compat 2020 vs 2022+ |
| `doc.Create.NewFloor(CurveArray, …)` | `Floor.Create(doc, IList<CurveLoop>, …)` | Create с 2022; NewFloor удалён 2023 | ветка compat 2020 vs 2022+ |
| `ElementId.IntegerValue` (Int32) | `ElementId.Value` (Int64), `ElementId(Int64)` | 2024 (deprecated) | хелперы element_id_value/make_element_id |
| IronPython-скрипты без шебанга | CPython через `#! python3` | политика проекта (locked) | конвенция + baseline для legacy |
| pythonnet 2.5 (мягкие касты int→enum) | pythonnet 3.x (строгие overload/enum, IList-маршалинг) | pyRevit 5.x | обходы уже в репо; консолидировать в compat |
| Только CLAUDE.md | AGENTS.md как кросс-агентный стандарт (60k+ репо) | 2025–2026 | D-05; Kilo нативно, Gemini import/настройка, Claude @import |
| `.claude/commands` отдельно от skills | commands merged into skills (равнозначны) | Claude Code, 2025–2026 | оба пути валидны; выбрать commands/ для изоляции от GSD |
| `.kilocode/workflows/` | `.kilo/commands/` (авто-миграция) | Kilo rebrand | адаптеры класть в новый путь, знать про legacy |

**Deprecated/устаревшее:** Revit 2025+ = .NET 8 (`NETCORE` в pyrevit.compat) — вне матрицы D-02, закладывается только расширяемостью compat (`/mm-new-compat`, deferred).

## Assumptions Log

| # | Claim | Section | Risk if Wrong |
|---|-------|---------|---------------|
| A1 | У пользователей pyRevit 5.x с pythonnet 3.x CPython-движком (вывод из коммитов 3fcf888/a93784a «pythonnet 3.x») — точная версия pyRevit не зафиксирована в репо | Standard Stack, Pitfalls 2/8 | Обходы compat остаются рабочими (каскадные), но приоритет веток может быть неоптимален; /mm-doctor стоит научить печатать версию pyRevit |
| A2 | Рекомендуемый размер icon.png ≈ 96×96 PNG (pyRevit масштабирует) | templates | Низкий: иконка отобразится и в других размерах; уточнить при первом скаффолде |
| A3 | pyRevit терпит BOM перед шебангом на текущей среде (кнопки с BOM исполняются CPython — косвенно по классу ошибок quick task 260709) | Baseline, Pitfall 1 | Если неверно и кнопки шли под IronPython — тем важнее правило MM003; поведение проверить на Revit UAT |
| A4 | `__revit__` доступен внутри lib-модулей на среде пользователей (работает в проде `ios_common_helpers.get_document`) | Pitfall 3 | compat должен использовать builtins-фолбэк — уже заложено в рекомендацию |
| A5 | Свежие Gemini CLI читают AGENTS.md нативно (agents.md указывает Gemini CLI в списке поддержки) | Pattern 1 | Не влияет: GEMINI.md-указатель с @AGENTS.md работает на всех версиях |

## Open Questions (RESOLVED)

1. **Судьба текущего содержимого root CLAUDE.md (GSD Release Map, Obsidian-поток)**
   - Известно: D-21 велит перенести graphify и Obsidian-поток в стандарт; D-05 — CLAUDE.md становится тонким указателем.
   - Неясно: GSD-блок («Синхронизируй gsd») — Claude-специфичен (GSD живёт в .claude) или общий? Рекомендация: GSD-блок оставить в CLAUDE.md под @AGENTS.md (Claude-специфика), Obsidian/graphify — в AGENTS.md. Решить при ревью AGENTS.md (гейт D-08-подобный).
   - **RESOLVED (план 03-05):** GSD-блок («Синхронизируй gsd») остаётся в CLAUDE.md под строкой `@AGENTS.md` как Claude-специфика; правила graphify и Obsidian-поток переносятся в AGENTS.md (D-21).
2. **Формат «полный документ конвенции»: AGENTS.md = весь текст vs AGENTS.md-ядро + docs/КОНВЕНЦИЯ.md**
   - Известно: D-05 называет AGENTS.md источником правды; spec agents.md не ограничивает размер, но контекст грузится каждую сессию каждым агентом.
   - Рекомендация: один AGENTS.md (~200–300 строк, только правила и ссылки на templates/tools); без отдельного дубля-документа. Если вырастет — выносить справочники (таблица ломающих изменений) в `agents/` с ссылками.
   - **RESOLVED (план 03-05):** один канонический AGENTS.md (~200–300 строк) — полный текст конвенции; отдельный docs/КОНВЕНЦИЯ.md не создаётся.
3. **`/mm-doctor`: как проверять «версию Revit vs поддерживаемые» без Revit**
   - Известно: из CLI Revit не запросить; установленные версии видны в реестре (`HKLM\SOFTWARE\Autodesk\Revit`) или по папкам `C:\Program Files\Autodesk\Revit 20XX`.
   - Рекомендация: doctor сверяет НАЙДЕННЫЕ установки с SUPPORTED_VERSIONS + проверяет целостность репо (vendored lib, полнота кнопок, орфаны layout, BOM-скан). Runtime-проверка версии остаётся в самих кнопках (D-03).
   - **RESOLVED (план 03-06, Task 3 /mm-doctor):** doctor сверяет установленные версии Revit (папки `C:\Program Files\Autodesk\Revit 20*` и/или реестр `HKLM\SOFTWARE\Autodesk\Revit`) с SUPPORTED_VERSIONS из revit_compat; runtime-проверку версии выполняют сами кнопки через require_supported_version (D-03).
4. **Адаптировать ли IFC_Двери/IFC_Окна в рамках фазы** — это исполнение `/mm-adopt-script` на реальных данных (отличный UAT), но не входит в объём ROADMAP. Рекомендация: UAT-сценарий фазы, отдельный quick task.
   - **RESOLVED:** адаптация IFC_Двери/IFC_Окна — UAT-сценарий фазы (03-VALIDATION.md §Manual-Only Verifications: прогон /mm-adopt-script на IFC_Двери), не обязательный объём; выполняется отдельным quick task при исполнении команды.

## Environment Availability

| Dependency | Required By | Available | Version | Fallback |
|------------|------------|-----------|---------|----------|
| Python (CPython) | tools/check_convention.py, unittest | ✓ | 3.14.2 (`python`), 3.13.7 (`python3`), py-launcher | — |
| git | /mm-save-session, /mm-update-repo | ✓ | 2.55.0.windows.3 | — |
| Node.js | GSD quick task tooling | ✓ | 20.20.0 | — |
| Claude Code (.claude инфраструктура) | адаптеры /mm-* | ✓ | GSD 60+ skills, hooks | — |
| Gemini CLI | адаптеры .gemini/commands | ✗ (не в PATH, нет .gemini/) | — | Файлы-адаптеры создать заранее; заработают при установке CLI |
| Kilo Code | адаптеры .kilo/commands, AGENTS.md | ✗ (нет .kilocode/.kilo) | — | Аналогично — адаптеры создаются «вперёд» |
| Autodesk Revit 2020/2022/2024 + pyRevit | UAT compat/шаблона | ✗ в headless-среде (прецедент: UAT вынесен в follow-up, STATE.md) | — | Manual UAT на рабочей станции (существующий процесс) |

**Missing dependencies with no fallback:** нет (блокеров исполнения фазы нет — все артефакты создаются локально).
**Missing dependencies with fallback:** Gemini CLI / Kilo Code (адаптеры пишутся заранее и проверяются форматно, не функционально); Revit UAT — ручной, вне headless.

## Validation Architecture

### Test Framework
| Property | Value |
|----------|-------|
| Framework | Python stdlib `unittest` (фреймворков в репо нет — подтверждено STACK.md и Glob; pytest сознательно не вводим) |
| Config file | none — Wave 0 создаёт `tools/tests/` |
| Quick run command | `py -3 -m unittest discover -s tools/tests -q` |
| Full suite command | `py -3 -m unittest discover -s tools/tests -v && py -3 tools/check_convention.py --all --baseline tools/convention_baseline.json` |

### Phase Requirements → Test Map
| Req ID | Behavior | Test Type | Automated Command | File Exists? |
|--------|----------|-----------|-------------------|-------------|
| CONV-CHECK | Чекер ловит каждое правило MM001–MM013 на bad-fixture и молчит на good-fixture | unit | `py -3 -m unittest tools.tests.test_check_convention -q` | ❌ Wave 0 |
| CONV-CHECK | Чекер по всему репо с baseline завершает exit 0 | integration | `py -3 tools/check_convention.py --all --baseline …` | ❌ Wave 0 |
| CONV-STD | templates/ проходит `--strict` на 100% | integration | `py -3 tools/check_convention.py "templates/НоваяКнопка.pushbutton" --strict` | ❌ Wave 0 |
| CONV-STD | revit_compat.py и шаблон синтаксически валидны | smoke | `py -3 -m py_compile "MM LAB.extension/lib/revit_compat.py" "templates/НоваяКнопка.pushbutton/script.py"` | ✅ (команда доступна сразу) |
| CONV-REG | MM007: layout-регистрация и орфаны детектируются | unit | входит в test_check_convention (fixtures с panel bundle.yaml) | ❌ Wave 0 |
| CONV-ADAPT | /mm-adopt-script: чекер→diff→approve→регистрация | manual-only | сценарий в Claude Code на IFC_Двери (ревью-гейт D-08 требует человека) | — |
| CONV-GSD | Quick task артефакты создаются (`.planning/quick/<id>/`, STATE-таблица) | manual/интеграция агента | проверка файлов после прогона команды | — |
| D-01..D-04 | compat-хелперы на Revit 2020/2022/2024 | manual-only (Revit smoke UAT — прецедент фазы 1) | чек-лист UAT в PLAN | — |

### Sampling Rate
- **Per task commit:** `py -3 -m unittest discover -s tools/tests -q` (< 10 s)
- **Per wave merge:** full suite command (unittest + чекер по репо + py_compile)
- **Phase gate:** full suite green + manual UAT чек-лист для compat/команд перед `/gsd-verify-work`

### Wave 0 Gaps
- [ ] `tools/tests/test_check_convention.py` — покрывает CONV-CHECK/CONV-REG (по одному тесту на правило)
- [ ] `tools/tests/fixtures/good_button/…` и `fixtures/bad_button/…` — эталонные деревья pushbutton (bad: BOM, wildcard, lookupParameter, нет README и т.д.)
- [ ] `tools/convention_baseline.json` — сгенерировать из фактического аудита 17 кнопок
- [ ] Framework install: не требуется (stdlib)

## Security Domain

Фаза — dev-tooling без сети/аутентификации/БД; применимые категории:

### Applicable ASVS Categories
| ASVS Category | Applies | Standard Control |
|---------------|---------|-----------------|
| V2 Authentication | no | — |
| V3 Session Management | no | — |
| V4 Access Control | частично | Ревью-гейт D-08: никакие правки/регистрации без явного одобрения человека; push только с подтверждением (D-18) |
| V5 Input Validation | yes | Сторонний скрипт — недоверенный ввод: анализ ТОЛЬКО `ast.parse` (без import/exec/eval чужого кода); пути аргументов нормализуются и ограничиваются репо |
| V6 Cryptography | no | — |

### Known Threat Patterns
| Pattern | STRIDE | Standard Mitigation |
|---------|--------|---------------------|
| Command injection через имена файлов (кириллица/пробелы/скобки в путях кнопок) в git/shell-вызовах mm-команд | Tampering | Всегда кавычить аргументы; предпочитать список-аргументы subprocess строковым шеллам; `git -c core.quotepath=false … -z` |
| Исполнение недоверенного стороннего скрипта при «проверке» | Elevation | Чекер статический (ast); запуск адаптированного кода — только человеком в Revit после ревью |
| Деструктивные git-операции | Repudiation/Tampering | `/mm-update-repo` — только fetch + ff-only pull при чистом дереве; никаких reset --hard/clean в командах |
| Тихая порча модели Revit шаблонным кодом | Tampering | Шаблон фиксирует Transaction+RollBack и верхнеуровневый показ ошибок (паттерн репо) |

## Sources

### Primary (HIGH confidence)
- Живой код репозитория: `MM LAB.extension/lib/ios_common_helpers.py`, `revit_ui_helpers.py`, скрипты «Мокрые зоны»/«Сброс потерь»/«Экспорт ПСО»/IFC_*, все bundle.yaml, `.planning/quick/260709-jko-*`, `.claude/skills/*` — прямое чтение
- [pyRevit docs (readthedocs): Extensions and Commands](https://pyrevit1.readthedocs.io/en/latest/creatingexts.html) — layout/`---`/`>>>`, bundle lib, .lib extensions
- [pyRevit forum: lib folder — builtin feature](https://discourse.pyrevitlabs.io/t/lib-folder-custom-modules-for-all-commands-in-the-extension/7764) — auto sys.path + reload
- [pyrevit compat.py (master, raw)](https://raw.githubusercontent.com/pyrevitlabs/pyRevit/master/pyrevitlib/pyrevit/compat.py) — NETCORE/get_elementid_value_func/_get_revit_version
- [Autodesk: What's New in the Revit 2024 API](https://blog.autodesk.io/whats-new-in-the-revit-2024-api/) + [Autodesk forum ElementId Int32/Int64](https://forums.autodesk.com/t5/revit-api-forum/revit-2024-elementid-integervalue-int32-vs-elementid-value-int64/td-p/11911934)
- [revitapidocs: ConvertFromInternalUnits(Double, ForgeTypeId) (2022)](https://www.revitapidocs.com/2022/60c6aac3-8306-c56e-b62f-b7011b9ad7b6.htm), [Floor.Create (2022)](https://www.revitapidocs.com/2022/a9c74a9f-46eb-a1b7-608e-2039f06be579.htm)
- [Claude Code docs: slash-commands/skills](https://code.claude.com/docs/en/slash-commands) — «commands merged into skills», frontmatter, $ARGUMENTS
- [Gemini CLI docs: custom commands](https://geminicli.com/docs/cli/custom-commands/), [GEMINI.md / contextFileName / @import](https://github.com/google-gemini/gemini-cli/blob/main/docs/cli/gemini-md.md)
- [Kilo docs: workflows/commands](https://kilo.ai/docs/customize/workflows), [kilo-docs agents-md](https://github.com/Kilo-Org/kilocode/blob/main/packages/kilo-docs/pages/customize/agents-md.md)
- [AGENTS.md spec](https://agents.md/)

### Secondary (MEDIUM confidence)
- [pyRevit forum: buttons showing despite not in bundle (6.4.0 launcher)](https://discourse.pyrevitlabs.io/t/buttons-showing-despite-not-being-in-bundle-pulldown-visibility-issue-in-latest-pyrevit/10078) — версионный дрейф layout
- [archi-lab: handling the Revit 2022 unit changes](https://archi-lab.net/handling-the-revit-2022-unit-changes/) — удаление DUT-API в 2022
- [learnrevitapi/форумы: `#! python3` включает CPython](https://www.learnrevitapi.com/newsletter/pyrevit-how-to-import-python3-packages-like-numpy-pandas-or-others)
- Гайды по мосту CLAUDE.md→AGENTS.md (`@AGENTS.md`, «Claude Code не читает AGENTS.md», май-2026 доки) — сходятся минимум 4 независимых источника
- Perplexity deep-research из CONTEXT.md (мультиверсия; agent-agnostic стандарт) — использованы как наводки, ключевые пункты перепроверены выше

### Tertiary (LOW confidence)
- Размер icon.png 96×96 — встречается в гайдах, официальную страницу не фиксировал → A2

## Metadata

**Confidence breakdown:**
- Механика pyRevit (lib auto-load, layout, шебанг): HIGH — доки + форум + подтверждение живым репо
- Ломающие изменения Revit API (Units/Floor/ElementId): HIGH — revitapidocs/Autodesk + pyrevit.compat
- Механизмы команд трёх агентов: HIGH (Claude/Gemini/Kilo — офиц. доки); нативный AGENTS.md у Gemini — MEDIUM (версионно)
- Чекер (ast, stdlib_module_names, baseline-подход): HIGH — stdlib, локальные Python 3.13/3.14 проверены
- Поведение среды пользователей (pyRevit версия, BOM, `__revit__` в lib): MEDIUM — только косвенные свидетельства репо (см. Assumptions A1/A3/A4)

**Research date:** 2026-07-24
**Valid until:** ~2026-08-24 (Revit API — стабильно; форматы команд агентов — быстро движутся, перепроверить пути `.kilo/` и Gemini contextFileName при исполнении)
