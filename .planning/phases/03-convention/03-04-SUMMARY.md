---
phase: 03-convention
plan: "04"
subsystem: tooling
tags: [pyrevit, python, template, baseline, convention, bundle-yaml]

requires:
  - phase: 03-convention (планы 03-01, 03-03)
    provides: чекер tools/check_convention.py (MM000–MM014, --strict/--baseline/--write-baseline, CLI-контракт)
  - phase: 03-convention (план 03-02)
    provides: MM LAB.extension/lib/revit_compat.py (require_supported_version, iter_count)
provides:
  - templates/НоваяКнопка.pushbutton/ — копируемый скелет кнопки (script.py + bundle.yaml + README.md), strict-чекер зелёный
  - Канонический lib-бутстрап (D-15) зафиксирован в шаблоне единственной формой
  - tools/convention_baseline.json — grandfathered legacy-кнопки (21 юнит, 202 нарушения из фактического аудита)
  - Чистый MM LAB.extension/MM Lab.tab/bundle.yaml — без орфана «ВОР» и хвостовых пробелов
  - Фазовый инвариант — py -3 tools/check_convention.py --all --baseline tools/convention_baseline.json -> exit 0
affects: [03-05 AGENTS.md, 03-06 mm-new-button, 03-06 mm-adopt-script, 03-06 mm-check, 03-07 адаптеры]

tech-stack:
  added: []
  patterns:
    - "Шаблон кнопки: шапка-канон -> канонический бутстрап -> revit_compat -> Transaction (Commit в try, RollBack+raise в except) -> TaskDialog, TODO-метки на местах правки (D-13)"
    - "Baseline-гейт: чекер полезен с первого дня — legacy заморожен, новый код проверяется строго"

key-files:
  created:
    - "templates/НоваяКнопка.pushbutton/script.py"
    - "templates/НоваяКнопка.pushbutton/bundle.yaml"
    - "templates/НоваяКнопка.pushbutton/README.md"
    - tools/convention_baseline.json
  modified:
    - "MM LAB.extension/MM Lab.tab/bundle.yaml"

key-decisions:
  - "Baseline генерируется только фактическим аудитом --write-baseline (не переписыванием цифр из RESEARCH): 21 юнит, 202 нарушения; записи снимаются при адаптации кнопки (/mm-adopt-script, план 03-06)"
  - "py_compile создаёт __pycache__ в папке кнопки и роняет следующий strict-прогон (MM013) — в verify-цепочках чистить __pycache__ после py_compile; git-гигиена уже покрыта .gitignore"
  - "REQUIREMENTS.md не трогаем: файл генерируется RELEASE_MAP/gsd_release_sync.py (T-xxx), CONV-* живут в ROADMAP.md; requirements-completed фиксируются во frontmatter SUMMARY (как в 03-01..03-03)"

patterns-established:
  - "Канонический lib-бутстрап: _SCRIPT_DIR -> 3 уровня «..» -> _EXTENSION_DIR -> _LIB_DIR -> sys.path.insert(0, _LIB_DIR); имя EXTENSION_ROOT и 4+ «..» запрещены (MM014)"
  - "Full suite фазы: unittest + --all --baseline + strict шаблона + py_compile (с очисткой __pycache__ после)"

requirements-completed: [CONV-STD, CONV-CHECK, CONV-REG]

coverage:
  - id: D1
    description: "Шаблон-скелет templates/НоваяКнопка.pushbutton: рабочая кнопка-пример с каноническим бутстрапом, revit_compat, транзакционным каркасом и TODO-метками; строгий чекер зелёный"
    requirement: CONV-STD
    verification:
      - kind: integration
        ref: "py -3 tools/check_convention.py \"templates/НоваяКнопка.pushbutton\" --strict"
        status: pass
      - kind: other
        ref: "py -3 -m py_compile \"templates/НоваяКнопка.pushbutton/script.py\""
        status: pass
    human_judgment: false
  - id: D2
    description: "Кнопка из шаблона реально запускается в Revit после копирования в панель (диалоги, транзакция, отчёт)"
    requirement: CONV-STD
    verification: []
    human_judgment: true
    rationale: "Revit API доступен только в среде Revit; headless-прогон невозможен. UAT-чек-лист — 03-VALIDATION.md §Manual-Only («Кнопка из шаблона видна на панели»)"
  - id: D3
    description: "Baseline legacy-кнопок из фактического аудита; полный прогон чекера по репо зелёный с baseline и красный без него"
    requirement: CONV-CHECK
    verification:
      - kind: integration
        ref: "py -3 tools/check_convention.py --all --baseline tools/convention_baseline.json (exit 0)"
        status: pass
      - kind: integration
        ref: "py -3 tools/check_convention.py --all (exit 1 — baseline фильтрует реальные нарушения)"
        status: pass
      - kind: unit
        ref: "py -3 -m unittest discover -s tools/tests -q (42 теста)"
        status: pass
    human_judgment: false
  - id: D4
    description: "Орфан «ВОР» удалён из tab bundle.yaml (папки ВОР.panel нет на диске), хвостовые пробелы убраны, живые записи панелей сохранены"
    requirement: CONV-REG
    verification:
      - kind: integration
        ref: "py -3 tools/check_convention.py --all --baseline tools/convention_baseline.json — MM007-орфанов нет"
        status: pass
      - kind: other
        ref: "grep-проверки: нет «ВОР», нет хвостовых пробелов, записи АРХИТЕКТУРА/ИОС/КООРДИНАЦИЯ и «-----» на месте"
        status: pass
    human_judgment: false

duration: 10min
completed: 2026-07-24
status: complete
---

# Phase 03 Plan 04: Шаблон кнопки + baseline конвенции Summary

**Копируемый скелет pushbutton в templates/ (канонический бутстрап + revit_compat + транзакционный каркас, strict-зелёный) и baseline из фактического аудита (21 юнит), сделавший полный прогон чекера по репо фазовым инвариантом**

## Performance

- **Duration:** 10 min
- **Started:** 2026-07-24T11:44:06Z
- **Completed:** 2026-07-24T11:53:51Z
- **Tasks:** 2
- **Files modified:** 5

## Accomplishments

- `templates/НоваяКнопка.pushbutton/` (вне MM Lab.tab — pyRevit его не грузит, D-14): рабочая кнопка-пример — канон шапки, канонический lib-бутстрап (D-15), `require_supported_version` первой строкой `main()` (D-03), пример чтения стен через `iter_count`, транзакционный каркас Commit/RollBack+raise, итоговый TaskDialog; 8 TODO-меток (D-13); `--strict` -> exit 0
- `tools/convention_baseline.json` сгенерирован фактическим аудитом `--write-baseline`: 21 юнит, 202 нарушения заморожены (включая нетронутые сторонние IFC_Двери/IFC_Окна); шаблон в baseline не попал
- Фазовый инвариант установлен: `--all --baseline` -> exit 0; `--all` без baseline -> exit 1 (baseline фильтрует реальные нарушения, а не пустой)
- Орфан «ВОР» удалён из tab bundle.yaml после проверки отсутствия папки `ВОР.panel` (T-03-11); живые записи АРХИТЕКТУРА/ИОС/КООРДИНАЦИЯ и разделитель «-----» сохранены; хвостовые пробелы убраны
- Полный интеграционный прогон (full suite из 03-VALIDATION.md) зелёный целиком: 42 unittest + --all --baseline + strict шаблона + py_compile compat и шаблона

## Task Commits

Each task was committed atomically:

1. **Task 1: Шаблон-скелет templates/НоваяКнопка.pushbutton** - `f370b05` (feat)
2. **Task 2: Правка tab bundle.yaml, генерация baseline, полный интеграционный прогон** - `effab8b` (chore)

## Files Created/Modified

- `templates/НоваяКнопка.pushbutton/script.py` - кнопка-пример: канон шапки, канонический бутстрап, revit_compat, Transaction, TaskDialog, 8 TODO (86 строк)
- `templates/НоваяКнопка.pushbutton/bundle.yaml` - каркас title/tooltip (ru+en_us) + author "GENPRO LAB" с TODO
- `templates/НоваяКнопка.pushbutton/README.md` - каркас разделов по образцу «Мокрых зон» + «Иконка» (icon.png ~96×96) + «Как использовать шаблон» (ручной путь и /mm-new-button)
- `tools/convention_baseline.json` - grandfathered-список: units {путь юнита -> коды правил}, generated, note
- `MM LAB.extension/MM Lab.tab/bundle.yaml` - удалён орфан «ВОР», убраны хвостовые пробелы и финальная строка из пробелов

## Decisions Made

- Baseline — только фактическим аудитом (репо менялось после RESEARCH): реальный аудит дал 21 юнит / 202 нарушения, в т.ч. вложенные `.vs/*.pushbutton`-папки мусора, которые rglob находит как юниты — заморожены как есть
- REQUIREMENTS.md не редактируется этим планом: файл принадлежит `RELEASE_MAP/gsd_release_sync.py` (проектный CLAUDE.md запрещает ручные правки), CONV-* требования отслеживаются в ROADMAP.md; `requirements mark-complete` пропущен — как в планах 03-01..03-03
- Импорты Autodesk в шаблоне — по одному на строку (стиль lib-модулей репо), а не в одну строку: оба варианта проходят чекер, выбран установленный паттерн

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 3 - Blocking] py_compile загрязняет папку шаблона и роняет следующий strict-прогон**
- **Found during:** Task 1 (verify-цепочка `py_compile && --strict`)
- **Issue:** `py -3 -m py_compile` создаёт `templates/НоваяКнопка.pushbutton/__pycache__/`, после чего `--strict` падает на MM013 (мусор в папке кнопки) — verify-команда сама себя блокирует при буквальном последовательном запуске
- **Fix:** в verify-прогонах после py_compile удаляется `__pycache__` (процедурная правка порядка запуска; код чекера не менялся — MM013 работает по замыслу, контракт CLI заморожен). Git-гигиена уже покрыта: `__pycache__/` в .gitignore
- **Files modified:** нет (изменение только процедуры запуска verify)
- **Verification:** повторный `--strict` после очистки -> exit 0; full suite (strict до py_compile) -> exit 0 целиком
- **Committed in:** — (файловых изменений нет)

---

**Total deviations:** 1 auto-fixed (1 blocking)
**Impact on plan:** Ноль изменений скоупа; эксплуатационная заметка зафиксирована в STATE.md (verify-цепочки чистят __pycache__ после py_compile).

## Known Stubs

TODO-метки в `script.py`/`bundle.yaml`/`README.md` шаблона — интенциональный дизайн (D-13: «минимальная рабочая кнопка-пример с явными TODO-метками в местах правки»), а не незавершённая работа: шаблон обязан содержать их как места правки для человека и `/mm-new-button`. Кнопка при этом полностью рабочая (читает и считает стены) и проходит строгий чекер.

## Issues Encountered

- Незакоммиченная правка рабочего дерева в tab bundle.yaml оказалась самим орфаном `- ВОР  ` (диф к HEAD) — конфликта с пользовательскими изменениями не возникло: условная правка плана (удалить только при отсутствии `ВОР.panel`) применена штатно, живые записи сохранены дословно
- В git-статусе остаются посторонние изменения не из этого плана: `.claude/settings.json`, `.planning/config.json` (правки пользователя/оркестратора), untracked `IFC_Двери.pushbutton`/`IFC_Окна.pushbutton` (живые сторонние скрипты — пойдут через `/mm-adopt-script`, план 03-06) и `docs/` — по scope boundary не тронуты и в коммиты не включены

## User Setup Required

None - no external service configuration required.

## Next Phase Readiness

- Волна 3 продолжается: план 03-05 (AGENTS.md + тонкие указатели) может ссылаться на шаблон как на канон бутстрапа и структуры кнопки
- Для 03-06: `/mm-new-button` копирует `templates/НоваяКнопка.pushbutton/`; `/mm-adopt-script` при адаптации кнопки удаляет её запись из `tools/convention_baseline.json` (механизм снятия T-03-12)
- Manual UAT (вне headless-среды): скопировать шаблон в панель, зарегистрировать в layout, pyRevit Reload — чек-лист в 03-VALIDATION.md §Manual-Only

## Self-Check: PASSED

- FOUND: templates/НоваяКнопка.pushbutton/script.py
- FOUND: templates/НоваяКнопка.pushbutton/bundle.yaml
- FOUND: templates/НоваяКнопка.pushbutton/README.md
- FOUND: tools/convention_baseline.json
- FOUND: коммит f370b05 (feat(03-04))
- FOUND: коммит effab8b (chore(03-04))
- PASS: full suite -> exit 0; --all без baseline -> exit 1

---
*Phase: 03-convention*
*Completed: 2026-07-24*
