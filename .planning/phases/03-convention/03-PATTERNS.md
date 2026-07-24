# Phase 3: convention - Pattern Map

**Mapped:** 2026-07-24
**Files analyzed:** ~30 новых/изменяемых файлов (сгруппированы в 12 кластеров)
**Analogs found:** 9 / 12 кластеров имеют прямой аналог в репо

## File Classification

| New/Modified File(s) | Role | Data Flow | Closest Analog | Match Quality |
|---|---|---|---|---|
| `AGENTS.md` (корень) | config/doc | — | корневой `CLAUDE.md` + `Мокрые зоны/README.md` (стиль RU-документации) | role-match |
| `CLAUDE.md` (правка → тонкий указатель) | config/doc | — | сам себя (сжимается) | exact |
| `GEMINI.md`, `.gemini/commands/mm-*.toml`, `.kilo/commands/mm-*.md`, `.kilocode/rules/00-mmlab.md` | config (адаптеры агентов) | — | нет в репо — образцы в RESEARCH §Code Examples | no-analog |
| `agents/commands/mm-*.md` (7 канонических процедур) | doc (процедура команды) | — | `.claude/skills/gsd-quick/SKILL.md` (структура), `.claude/skills/graphify/SKILL.md` (Usage-блок) | role-match |
| `.claude/commands/mm-*.md` (7 адаптеров) | config (command) | — | `.claude/skills/gsd-quick/SKILL.md` (frontmatter) | role-match |
| `MM LAB.extension/lib/revit_compat.py` | utility (shared lib) | request-response (API-обёртки) | `MM LAB.extension/lib/ios_common_helpers.py` | exact |
| `tools/check_convention.py` | utility (CLI-чекер) | file-I/O + transform (AST) | нет CLI-инструментов в репо; ядро — RESEARCH §Code Examples (AST) | no-analog |
| `tools/convention_baseline.json` | config (data) | file-I/O | нет — плоский JSON, схему задаёт чекер | no-analog |
| `tools/tests/*` + fixtures | test | batch | нет тестов в репо — stdlib `unittest` (RESEARCH §Validation) | no-analog |
| `templates/НоваяКнопка.pushbutton/script.py` | pushbutton script (шаблон) | request-response (Revit UI) | `Сброс потерь/script.py` (структура main) + `Мокрые зоны/script.py` (шапка) | exact |
| `templates/НоваяКнопка.pushbutton/bundle.yaml`, `README.md` | config/doc | — | `Мокрые зоны.pushbutton/bundle.yaml` + `README.md` | exact |
| `MM LAB.extension/MM Lab.tab/bundle.yaml` (правка: убрать орфан `ВОР`, хвостовые пробелы) | config | — | `АРХИТЕКТУРА.panel/bundle.yaml` | exact |

## Pattern Assignments

### `MM LAB.extension/lib/revit_compat.py` (utility, shared lib)

**Analog:** `MM LAB.extension/lib/ios_common_helpers.py` — единственный полноценный shared-модуль; compat встаёт рядом тем же механизмом (pyRevit авто-добавляет `MM LAB.extension/lib` в sys.path).

**Шапка + импорты** (`ios_common_helpers.py` строки 1–23):
```python
#! python3
# -*- coding: utf-8 -*-

import math

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import BuiltInParameter
from Autodesk.Revit.DB import ElementId
...  # по одному импорту на строку, отсортированы
from Autodesk.Revit.UI import TaskDialog
```
Порядок: stdlib → `clr.AddReference` → `Autodesk.Revit.*` (алфавитно, по одному на строку). У compat добавить docstring по канону шапки (см. шаблон ниже) — в `ios_common_helpers` его нет, это дефект, не паттерн.

**ElementId Int64-хелпер — уже существует, НЕ дублировать, а перенести/делегировать** (`ios_common_helpers.py` строки 74–81):
```python
def element_id_value(element_id):
    if element_id is None:
        return -1
    try:
        return element_id.Value          # Revit 2024+: Int64
    except Exception:
        return element_id.IntegerValue   # 2020/2022
```

**Обход №1 get_Parameter — BIP→имя→LookupParameter с кешем** (`ios_common_helpers.py` строки 145–182):
```python
_BIP_NAME_CACHE = {}

def _bip_to_lookup_name(document, built_in_parameter):
    bip_int = int(built_in_parameter)
    if bip_int not in _BIP_NAME_CACHE:
        name = None
        param_element = document.GetElement(ElementId(bip_int))
        if param_element is not None:
            try:
                name = param_element.GetDefinition().Name
            except Exception:
                name = None
        _BIP_NAME_CACHE[bip_int] = name
    return _BIP_NAME_CACHE[bip_int]

def get_parameter(element, built_in_parameter, *names):
    name = _bip_to_lookup_name(element.Document, built_in_parameter)
    if name:
        parameter = element.LookupParameter(name)
        if parameter is not None:
            return parameter
    return get_parameter_by_names(element, *names)   # fallback по локализованным именам
```

**Обход №2 get_Parameter — pythonnet `__overloads__`** (`Мокрые зоны/script.py` строки 80–96):
```python
def _get_param(element, bip):
    try:
        return element.get_Parameter.__overloads__[BuiltInParameter](bip)
    except (TypeError, AttributeError, KeyError):
        return element.get_Parameter(bip)
```
**Требование D-01/anti-pattern:** compat выбирает ОДИН канонический каскад из этих двух (рекомендация RESEARCH: `__overloads__` → BIP→имя→LookupParameter → fallback-имена), не создаёт третий вариант.

**UI-репортинг ошибок** (`ios_common_helpers.py` строки 92–102):
```python
def show_error(command_name, ex):
    TaskDialog.Show(command_name, u"Ошибка:\n{0}".format(to_text(ex)))

def get_document(command_name):
    uidoc = __revit__.ActiveUIDocument
    if uidoc is None or uidoc.Document is None:
        TaskDialog.Show(command_name, u"Открой проект Revit и повтори команду.")
        return None
    return uidoc.Document
```
Внимание (Pitfall 3 из RESEARCH): `__revit__` в lib-модуле — негарантированный контракт; в compat детекцию версии делать каскадом с `getattr(builtins, "__revit__", None)`. `require_supported_version()` (D-03) следует паттерну `get_document`: TaskDialog с понятным русским текстом + мягкий выход.

Скелет модуля (SUPPORTED_VERSIONS, units-ветка 2020 vs 2022+, `create_floor`, `to_net_list`, `enum_from_int`) — см. RESEARCH §Pattern 3 и §Code Examples (проверенные сигнатуры API).

---

### `templates/НоваяКнопка.pushbutton/script.py` (pushbutton script, request-response)

**Шапка — канон** (`Мокрые зоны/script.py` строки 1–26):
```python
#! python3
# -*- coding: utf-8 -*-
"""Название кнопки

Что делает, кратко (несколько строк).

Совместимость: Revit 2020 / 2022 / 2024
Зависимости: нет
"""

__title__ = "Мокрые\nзоны"
__author__ = "GENPRO LAB"
__doc__ = ("Однострочное описание для tooltip.")
```

**lib-бутстрап — канонизировать по D-15.** В репо два конфликтующих варианта:
- `Мокрые зоны/script.py` строки 33–37: 4×`..` → корень РЕПО → vendored `lib/`, `insert(0)` (имя `EXTENSION_ROOT` лжёт);
- `Сброс потерь/script.py` строки 12–16: 3×`dirname` → `MM LAB.extension` → first-party `lib/`, `append`.

Канон для шаблона (форма «Мокрых зон», путь «Сброса потерь», честные имена — RESEARCH §Pattern 4):
```python
import os, sys
_SCRIPT_DIR = os.path.dirname(__file__)
# pushbutton → panel → tab → MM LAB.extension
_EXTENSION_DIR = os.path.normpath(os.path.join(_SCRIPT_DIR, "..", "..", ".."))
_LIB_DIR = os.path.join(_EXTENSION_DIR, "lib")
if os.path.isdir(_LIB_DIR) and _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)
```

**Структура тела — по «Сбросу потерь»** (`Сброс потерь/script.py` строки 29–82) — самая чистая кнопка репо:
```python
COMMAND_NAME = u"Сброс потерь"

try:
    doc = get_document(COMMAND_NAME)
    if doc:
        ...
        transaction = Transaction(doc, COMMAND_NAME)
        transaction.Start()
        try:
            for element in targets:
                ...
            transaction.Commit()
        except Exception:
            transaction.RollBack()
            raise

        TaskDialog.Show(
            COMMAND_NAME,
            u"Переведено: {0}\nУже были: {1}\nПропущено: {2}".format(...),
        )
except Exception as ex:
    show_error(COMMAND_NAME, ex)
```
Паттерн: `Transaction.Start()`/`Commit()` в try, `RollBack()` + `raise` в except, верхнеуровневый `show_error`, итоговый `TaskDialog` со счётчиками. В шаблон добавить `require_supported_version(COMMAND_NAME)` из compat перед работой (D-03) и `# TODO`-метки (D-13). Оборачивание в `main()` + `if __name__ == "__main__"` — как в «Мокрых зонах» (строки 674–687), но БЕЗ голого `except:` (там дефект MM011; в шаблоне — `except Exception`).

---

### `templates/НоваяКнопка.pushbutton/bundle.yaml` (config)

**Analog:** `Мокрые зоны.pushbutton/bundle.yaml` (целиком, 7 строк):
```yaml
title:
  en_us: "Wet\nZones"
  ru: "Мокрые\nзоны"
tooltip:
  en_us: "Check rooms for intersection with wet room projections from the level above"
  ru: "Проверка помещений на пересечение с проекцией мокрых помещений уровнем выше"
author: "GENPRO LAB"
```
Копировать структуру дословно, значения → `# TODO`.

---

### `templates/НоваяКнопка.pushbutton/README.md` (doc)

**Analog:** `Мокрые зоны.pushbutton/README.md` — русский, разделы: `# Название`, `## Описание`, `## Логика работы` (нумерованный список), `## Параметры помещений` (таблица используемых параметров Revit), `## Зависимости`, `## Совместимость`. Каркас README шаблона = эти заголовки с TODO.

---

### Правки `bundle.yaml` (tab + panel) и правило регистрации для `/mm-adopt-script`

**Analog:** `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/bundle.yaml`:
```yaml
layout:
  - Мокрые зоны
  - Экспорт ПСО
  - Кривые панели
  ...
```
Имена БЕЗ суффикса `.pushbutton`, порядок = порядок кнопок. Команда приёмки дописывает строку сюда (D-11). Дефекты для исправления/детекции в tab `bundle.yaml` (строки 4–9): орфан `ВОР` (панели `ВОР.panel` нет) и хвостовые пробелы (`- ВОР  `, пустая строка `  `) — тест-кейсы правила MM007.

---

### `agents/commands/mm-*.md` и `.claude/commands/mm-*.md` (агентские команды)

**Analog frontmatter:** `.claude/skills/gsd-quick/SKILL.md` строки 1–13:
```yaml
---
name: gsd-quick
description: "Execute a quick task with GSD guarantees (atomic commits, state tracking) but skip optional agents"
argument-hint: "[--full] [--validate] [--discuss] [--research]"
allowed-tools:
  - Read
  - Write
  - Edit
  - Bash
---
```
**Analog структуры процедуры:** `.claude/skills/graphify/SKILL.md` — заголовок `# /command`, абзац назначения, блок `## Usage` с вариантами вызова. Канонические процедуры в `agents/commands/` — на русском, с гейтами D-08/D-10/D-17/D-18 текстом; адаптеры `.claude/commands/mm-*.md` — frontmatter + 2–4 строки «Прочитай `agents/commands/mm-X.md` и выполни. Аргументы: $ARGUMENTS». Адаптеры Gemini (`.gemini/commands/mm-check.toml`) и Kilo (`.kilo/commands/mm-check.md`) — дословно из RESEARCH §Code Examples (в репо аналогов нет). Класть в `.claude/commands/`, НЕ в `.claude/skills/` (там GSD-манифест).

**Шаблон коммита `/mm-save-session`** — утверждён в CONTEXT §specifics (обязательный, копировать дословно); прежний вариант — корневой `CLAUDE.md` §Obsidian. Git: пофайловый стейджинг, `git -c core.quotepath=false status --porcelain` (кириллические пути).

---

### Quick task из `/mm-adopt-script` (D-09)

**Analog:** `.planning/quick/260709-jko-get-parameter-typeerror-builtinparameter/260709-jko-PLAN.md`:
```markdown
---
quick_id: 260709-jko
status: complete
---

# Quick Task 260709-jko: <название>

## Task
## Root cause / Fix
```
Артефакты: папка `.planning/quick/<id>-<slug>/` с `<id>-PLAN.md` + `<id>-SUMMARY.md` + строка в таблице STATE.md. В Claude Code — делегировать в `/gsd-quick`; для Gemini/Kilo процедура описывает создание файлов вручную по этому образцу.

---

### `AGENTS.md` + тонкие указатели

**Analog содержимого:** корневой `CLAUDE.md` (graphify-правило, Obsidian-поток, шаблон коммита — переносятся по D-21; GSD Release Map — остаётся Claude-специфичным в CLAUDE.md, Open Question Q1). Стиль prose — русский, как README «Мокрых зон». Новый `CLAUDE.md`: первая строка `@AGENTS.md`, ниже — только Claude-специфика (GSD). `GEMINI.md`: `@AGENTS.md`. Kilo читает AGENTS.md нативно; опционально `.kilocode/rules/00-mmlab.md` в одну строку.

---

### `tools/check_convention.py`, `tools/tests/*`, `convention_baseline.json`

Аналогов в репо нет (первый CLI/тест-инструмент). Каркас — RESEARCH §Pattern 5 (правила MM001–MM013, `--strict`/`--json`/`--baseline`, exit-коды, `sys.stdout.reconfigure(encoding="utf-8")`) и §Code Examples (AST-ядро MM008/MM009/MM011). Тесты — stdlib `unittest`, discover из `tools/tests`. Fixtures: `good_button/` строится из шаблона; `bad_button/` — по реальным дефектам `IFC_Двери.pushbutton` (нет шебанга, `from Autodesk.Revit.DB import *`, `LookupParameter("GP_23_Назначение")`, голые `except: pass`, нет bundle.yaml/README) и BOM-кнопок ИОС. Baseline генерировать из фактического аудита 17 кнопок (RESEARCH §Baseline Audit).

## Shared Patterns

### Обязательная шапка скрипта
**Source:** `Мокрые зоны/script.py` строки 1–26 (см. выше). **Apply to:** шаблон, `revit_compat.py`, все примеры в AGENTS.md, правила MM001–MM004 чекера.

### Транзакция + верхнеуровневая ошибка
**Source:** `Сброс потерь/script.py` строки 53–82 (см. выше). **Apply to:** шаблон script.py, раздел «Транзакции» в AGENTS.md.

### UI-alert
**Source:** `MM LAB.extension/lib/revit_ui_helpers.py` строки 12–16:
```python
def alert(message, title=u"Сообщение"):
    dialog = TaskDialog(title)
    dialog.MainContent = message
    dialog.CommonButtons = TaskDialogCommonButtons.Ok
    dialog.Show()
```
**Apply to:** шаблон (импорт из lib), fail-fast сообщение compat.

### Стиль импортов lib-модулей
**Source:** `ios_common_helpers.py` строки 1–23. **Apply to:** `revit_compat.py`, шаблон, правило MM008 (белый список: stdlib + `clr`,`System`,`Autodesk`,`pyrevit` + `MM LAB.extension/lib/*.py` + vendored `openpyxl`,`et_xmlfile`).

## No Analog Found

| File | Role | Data Flow | Reason / что использовать |
|------|------|-----------|--------|
| `tools/check_convention.py` + tests + baseline | utility/test | file-I/O, transform | В репо нет CLI-инструментов и тестов → RESEARCH §Pattern 5, §Validation Architecture |
| `.gemini/commands/*.toml`, `.kilo/commands/*.md`, `GEMINI.md` | config | — | Gemini/Kilo не установлены → готовые образцы в RESEARCH §Code Examples |
| `templates/…/icon.png` | asset | — | Иконки в репо разнобойные (`icons8-дверь-100.png`) → плейсхолдер + инструкция, ~96×96 PNG (RESEARCH A2) |

## Metadata

**Analog search scope:** `MM LAB.extension/lib/`, `MM LAB.extension/MM Lab.tab/**` (панели АРХИТЕКТУРА/ИОС), `.claude/skills/`, `.planning/quick/`, корневые `CLAUDE.md`/`bundle.yaml`
**Files scanned:** ~15 прочитано напрямую (lib-модули, 2 эталонных script.py, bundle.yaml×3, README, SKILL.md×2, quick-task PLAN)
**Pattern extraction date:** 2026-07-24
